import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

var spEventsParser = {
    parseEvents: function (events, start, end) {
        var full = [];
        for (var i = 0; i < events.length; i++) {
            full = full.concat(this.parseEvent(events[i], start, end));
        }
        return full;
    }
}

const getEvents = () => {
    console.log("Context : ", this.context);
    let eventDetails = this.props.context.spHttpClient.get(`https://technovert2020.sharepoint.com/sites/Technovert/_api/lists/GetByTitle('SampleEvents')/items`,
        SPHttpClient.configurations.v1)
        .then(async (response) => {
            const responseJSON = await response.json();
            console.log(responseJSON);
            let parsedArray = spEventsParser.parseEvent(responseJSON);
            console.log("Parsed Array : ", parsedArray);
            return responseJSON;
        });

    console.log("Event Details : ", eventDetails);
}