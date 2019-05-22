import * as express from 'express'
import * as url from 'url'
import * as request from 'request'
import * as cookieParser from 'cookie-parser'
import * as moment from 'moment'
import * as bodyParser from 'body-parser'

const app = express()
app.set('view engine', 'ejs')
app.use(cookieParser())
app.use(bodyParser.urlencoded({extended: true}))

const TENANT_ID = process.env.TENANT_ID
const AUTH_BASE_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`
const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET

app.get('/', (req, res) => {
    const accessToken = req.cookies.accessToken
    if (accessToken) {
        renderForm(req, res)
    } else {
        renderLoginPage(req, res)
    }
})

const renderForm = (req: express.Request, res: express.Response): void => {
    const accessToken = req.cookies.accessToken
    request({
        url: 'https://graph.microsoft.com/v1.0/me/calendar/calendarView',
        qs: {
            'startDateTime': new Date().toISOString(),
            'endDateTime': new Date(Date.now() + 604800000).toISOString(),
            '$orderby': 'start/datetime'
        },
        method: 'GET',
        json: true,
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Prefer': 'outlook.timezone="Asia/Tokyo"'
        }
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.render('form.ejs', {
                events: body.value.map(event => {
                    return {
                        start: moment(event.start.dateTime).format('YYYY/MM/DD HH:mm'),
                        subject: event.subject,
                        location: event.location.displayName
                    }
                })
            })
        }
    })
}

const renderLoginPage = (req: express.Request, res: express.Response): void => {
    const endpointUrl = `${AUTH_BASE_URL}/authorize`
    const queryParams = {
        'client_id': CLIENT_ID,
        'response_type': 'code',
        'redirect_uri': `${process.env.APP_URL}/callback`,
        'scope': 'User.Read Calendars.ReadWrite.Shared',
        'prompt': 'consent'
    }
    const targetUrl = url.parse(endpointUrl, true)
    const query = targetUrl.query
    Object.keys(queryParams).forEach(key => {
        query[key] = queryParams[key]
    })
    res.render('login.ejs', {
        authorizationUrl: url.format(targetUrl)
    })
}

app.get('/callback', (req, res) => {
    request({
        url: `${AUTH_BASE_URL}/token`,
        method: 'POST',
        form: {
            'grant_type': 'authorization_code',
            'client_id': CLIENT_ID,
            'code': req.query.code,
            'redirect_uri': `${process.env.APP_URL}/callback`,
            'client_secret': CLIENT_SECRET
        },
        json: true
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.cookie('accessToken', body.access_token, {
                maxAge: 60 * 60 * 1000,
                httpOnly: false
            })
            res.redirect('/')
        }
    })
})

app.post('/', (req, res) => {
    const accessToken = req.cookies.accessToken
    const subject = req.body.subject
    const [locationEmailAddress, locationName] = req.body.location.split('/')
    const startDate = req.body.startDate.split('/')
    const startTime = req.body.startTime.split(':')
    const startDateTime = moment([
        Number(startDate[0]), Number(startDate[1]) - 1, Number(startDate[2]),
        Number(startTime[0]), Number(startTime[1]), 0])
    const length = req.body.length
    const endDateTime = moment(startDateTime).add(length, 'm')
    request({
        url: 'https://graph.microsoft.com/v1.0/me/calendar/events',
        method: 'POST',
        json: {
            subject,
            start: {
                dateTime: startDateTime.format('YYYY-MM-DDTHH:mm:ss'),
                timeZone: 'Asia/Tokyo'
            },
            end: {
                dateTime: endDateTime.format('YYYY-MM-DDTHH:mm:ss'),
                timeZone: 'Asia/Tokyo'
            },
            location: {
                locationEmailAddress,
                displayName: locationName
            },
            attendees: [
                {
                    emailAddress: {
                        address: locationEmailAddress
                    },
                    type: 'resource'
                }
            ]
        },
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Prefer': 'outlook.timezone="Asia/Tokyo"'
        }
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.redirect('/')
        }
    })
})

app.listen(process.env.PORT || 1337)
