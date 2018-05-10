import * as React from 'react';
import { IReminderEventsProps } from './IReminderEventsProps';
import { SPHttpClient } from '@microsoft/sp-http';
import helper from "./EventsHelper";
import * as moment from 'moment';

export interface IReminderEventsState {
    results?: any[][];
    totalRowsCount: any[][];
    error: any;
}

export default class ReminderEventsDev extends React.Component<IReminderEventsProps, IReminderEventsState> {
    private _helper: helper;
    public constructor() {
        super();
        this._helper = new helper();
        this.state = {
            results: [],
            totalRowsCount: [],
            error: ""
        };
    }

    public componentDidMount() {
        try {
            this.InitializeSearchState();
        }
        catch (exception) {
            this.setState({ error: exception.message });
        }
    }

    public componentWillReceiveProps(): void {
        try {
            this.InitializeSearchState();
        }
        catch (exception) {
            this.setState({ error: exception.message });
        }
    }

    public InitializeSearchState() {
        try {
            let pathFilter: any = "";
            let searchquery: string = 'ContentType:"OneCMS_Event" AND RefinableString28:"Reminder"';
            if (window.location.href.indexOf('o365sitename') != -1) {
                searchquery = 'ContentType:"OneCMS_Event" AND RefinableString27:"Reminder"';
            }
            if (this.props.WPProperties.EventType == "All") {
                searchquery = 'ContentType:"OneCMS_Event"';
            }
            let selectProperties: string = "Title,path,ListItemId,EventDateOWSDATE,EndDateOWSDATE,LocationOWSTEXT,Location,Abstract,RefinableString27,RefinableString28";
            let NumberOfItems = this.props.WPProperties.NumberOfItems === undefined ? 4 : this.props.WPProperties.NumberOfItems;
            let itemSorting: string = "RefinableDate03:descending";
            let siteInfoJSON;
            let siteName = this.props.context.pageContext.web.absoluteUrl.slice(this.props.context.pageContext.web.absoluteUrl.lastIndexOf("/") + 1);
            // If ContentFlow else ThisSite
            if (!this._helper.isNullOrUndefinedOrEmpty(JSON.parse(sessionStorage.getItem("SiteInfo")))) {
                siteInfoJSON = JSON.parse(sessionStorage.getItem("SiteInfo"))[siteName];
                if (!this._helper.isNullOrUndefinedOrEmpty(siteInfoJSON) && !this._helper.isNullOrUndefinedOrEmpty(siteInfoJSON.PathFilter.trim()))
                    pathFilter = " " + siteInfoJSON.PathFilter.trim();
                else
                    pathFilter = "";
            }
            else
                pathFilter = ` U:"${this.props.context.pageContext.web.serverRelativeUrl}"`;
            searchquery += pathFilter.replace(/\-/g, "/-").replace(-[0 - 9], "/$&");
            let rowLimit = this.props.WPProperties.NumberOfItems == undefined ? 4 : this.props.WPProperties.NumberOfItems;
            let searchState = {};
            searchState["context"] = this.props.context;
            searchState["query"] = searchquery;
            searchState["fields"] = selectProperties;
            searchState["rowPerPage"] = rowLimit;
            searchState["maxResults"] = rowLimit;
            searchState["sorting"] = itemSorting;
            this.GetSearchResults(searchState);

        }
        catch (exception) {
            this.setState({ error: exception.message });
        }
    }

    public render(): React.ReactElement<IReminderEventsProps> {
        /********************************************* START : View all Show/Hide Logic ******************************************** */
        // Loop through results to get number of itemd selected in the widget
        let arrResults: any[] = [];
        let bViewall: boolean = false;

        let totalRowCount = 0;

        let tempResults = [];
        if (this.props.InstanceId == undefined)
            tempResults = this.state.results;
        else {
            tempResults = this.state.results[this.props.InstanceId];
            totalRowCount = this.state.totalRowsCount[this.props.InstanceId];
            if (totalRowCount >= parseInt(this.props.WPProperties.NumberOfItems))
                bViewall = true; //Reminder Listing condition needs tobe implemented
        }

        if (tempResults != undefined && tempResults.length > 0) {
            for (let i = 0; i < tempResults.length; i++) {
                arrResults[i] = tempResults[i];
                if (i == (this.props.WPProperties.NumberOfItems === undefined ? 3 : (this.props.WPProperties.NumberOfItems - 1)))
                    break;
            }
        }

        let boxBorderClassName = "";
        if (this.props.WPProperties.IsBoxed == "Yes") {
            boxBorderClassName = " oc-widget-level-box-model";
        }

        let cardViewClass = "oc-w-row listview oc-no-separator";
        if (this.props.WPProperties.NumberOfItems >= 4) {
            cardViewClass = "oc-w-row listview oc-no-separator";
        }
        else if (this.props.WPProperties.NumberOfItems == 3) {
            cardViewClass = "oc-w-row listview oc-no-separator oc-items-count-3";
        }
        else if (this.props.WPProperties.NumberOfItems == 2) {
            cardViewClass = "oc-w-row listview oc-no-separator oc-items-count-2";
        }
        else if (this.props.WPProperties.NumberOfItems == 1) {
            cardViewClass = "oc-w-row listview oc-no-separator oc-items-count-1";
        }
        let ListingPath = this.props.context.pageContext.web.absoluteUrl + "/Sitepages/listing.aspx/Events";
        let DetailPath = this.props.context.pageContext.web.absoluteUrl + "/Sitepages/Detail.aspx/Events/mode=view";
        return (
            <div className={bViewall ? boxBorderClassName : boxBorderClassName + " oc-no-viewall"}>
                <div id="widget-events">
                    {this.state.error.length > 0 ? "Something went wrong. Webpart ERROR :" + this.state.error :
                        arrResults != undefined && arrResults.length > 0 ?
                            <div id="events_widget">
                                <section className="oc-widget-section">
                                    <div className="oc-container">
                                        <div className="oc-widget-title">
                                            <div>{this.props.WPProperties.WebPartTitle === undefined || this.props.WPProperties.WebPartTitle === null || this.props.WPProperties.WebPartTitle === "" ? this.props.WPProperties.description : this.props.WPProperties.WebPartTitle}</div>
                                            <div className="oc-viewall-top">
                                                {bViewall ?
                                                    <a data-parent="events" href={ListingPath} className="oc-btn-viewall">
                                                        View All <i className="arr"></i>
                                                    </a>
                                                    : null
                                                }
                                            </div>
                                        </div>
                                        <div className="oc-widget-content oc-hr-reminders">
                                            <div className={cardViewClass}>
                                                {arrResults != undefined && arrResults.length > 0 ?
                                                    arrResults.map((item, i) => {
                                                        let eventStartDate = this._helper.formatDate(item.EventDateOWSDATE);
                                                        let eventStartTime = this._helper.formatAMPM(item.EventDateOWSDATE);
                                                        let eventEndDate = this._helper.formatDate(item.EndDateOWSDATE);
                                                        let eventEndTime = this._helper.formatAMPM(item.EndDateOWSDATE);
                                                        let a = moment(item.EventDateOWSDATE);
                                                        let b = moment();
                                                        let daysRem = a.diff(b, 'days');
                                                        return (
                                                            <div className="oc-w-col oc-apply-block">
                                                                <div className="oc-widget-box oc-event-widget">
                                                                    <div className="oc-date-txt">
                                                                        <span className="oc-event-date">
                                                                            {eventEndDate}
                                                                            {String(daysRem).indexOf('-') >= 0 ?
                                                                                <span className="oc-time-txt">
                                                                                    {daysRem ? String(daysRem).replace('-', '') + " days ago" : ""}
                                                                                </span> :
                                                                                <span className="oc-time-txt">
                                                                                    {daysRem > 1 ? daysRem + " days remaining" : daysRem + " day remaining"}
                                                                                </span>
                                                                            }
                                                                        </span>
                                                                    </div>
                                                                    <div className="oc-content">
                                                                        <a data-parent="events" href={DetailPath + "?_Id=" + item.ListItemId + "&SiteURL=/sites/" + item.path.toLowerCase().split("/lists/")[0].split("/sites/")[1]} data-src={DetailPath + "?_Id=" + item.ListItemId + "&SiteURL=/sites/" + item.path.toLowerCase().split("/lists/")[0].split("/sites/")[1]} className="oc-link-txt andes oc-lines2">
                                                                            {this._helper.GetTrimmedText(item.Title, 65) }</a>
                                                                    </div>
                                                                    <div className="oc-event-desc ms-hide">
                                                                        {!this._helper.isEmptyString(item.Abstract) ? this._helper.GetTrimmedText(item.Abstract, 65) : ''}
                                                                    </div>
                                                                    <div className="oc-event-location ms-hide">
                                                                        {item.Location ? <span className="oc-eve-loc">Location: {this._helper.getTermString(item.Location) }&nbsp; </span> : null }
                                                                        {item.LocationOWSTEXT ? <span className="oc-eve-room">Room Number: {item.LocationOWSTEXT}</span> : null }
                                                                    </div>
                                                                </div>
                                                            </div>);
                                                    }

                                                    ) : null}
                                            </div>
                                        </div>
                                        {bViewall ?
                                            <div className="oc-viewall-bottom">
                                                <a data-parent="events" href={ListingPath} className="oc-btn-viewall">
                                                    View All <i className="arr"></i>
                                                </a>
                                            </div>
                                            : null
                                        }
                                    </div>
                                </section>
                            </div>
                            :
                            <div>
                                <section className="oc-widget-section">
                                    <div className="oc-container">
                                        <div className="oc-widget-title">
                                            <div>{this._helper.isNullOrUndefinedOrEmpty(this.props.WPProperties.WebPartTitle)
                                                ? this.props.WPProperties.description : this.props.WPProperties.WebPartTitle}</div>
                                        </div>
                                        <div className="oc-widget-content"><div className="oc-no-items">No items to show</div></div>
                                    </div>
                                </section>
                            </div>

                    }
                </div>
            </div>
        );
    }

    public GetSearchResults(searchState) {
        let that = this;
        let url: string = searchState.context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=";
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.query) ? `'${searchState.query}'` : "'*'";
        // Check if there are fields provided
        url += '&selectproperties=';
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.fields) ? `'${searchState.fields}'` : "'path,title'";
        url += "&startrow=";
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.startRow) ? searchState.startRow : 0;
        url += "&rowsperpage=";
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.rowPerPage) ? searchState.rowPerPage : 12;
        // Add the rowlimit
        url += "&rowlimit=";
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.maxResults) ? searchState.maxResults : 12;
        // Add sorting
        url += !this._helper.isNullOrUndefinedOrEmpty(searchState.sorting) ? `&sortlist='${searchState.sorting}'` : "";
        // Add the client type
        url += "&clienttype='ContentSearchRegular'";
        url += '&trimduplicates=false';

        this.GetSearchData(searchState.context, url).then((res: any) => {
            let resultsRetrieved = false;
            if (res !== null) {
                if (typeof res.PrimaryQueryResult !== 'undefined') {
                    if (typeof res.PrimaryQueryResult.RelevantResults !== 'undefined') {
                        if (typeof res.PrimaryQueryResult.RelevantResults !== 'undefined') {
                            if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {
                                if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {
                                    that.setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, res.PrimaryQueryResult.RelevantResults.TotalRows, searchState.fields, searchState.context);
                                }
                            }
                        }
                    }
                }
            }
        });
    }

    public setSearchResults(crntResults: any, totalRowsCount: number, fields: string, context: any): void {
        let _results = [];
        let _totalRows = [];
        if (crntResults.length > 0) {
            let flds: string[] = fields.toLowerCase().split(',');
            let temp: any[] = [];
            let tempResults: any[][] = [];
            crntResults.forEach((result) => {
                // Create a temp value
                let val: Object = {};
                result.Cells.forEach((cell: any) => {
                    if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {
                        // Add key and value to temp value
                        val[cell.Key] = cell.Value;
                    }
                });
                // Push this to the temp array
                temp.push(val);
            });
            _results[context.instanceId] = temp;
            _totalRows[context.instanceId] = totalRowsCount;
        }
        this.setState({
            results: _results,
            totalRowsCount: _totalRows
        });
    }

    public GetSearchData(context: any, url: string): Promise<any> {
        return context.spHttpClient.get(url, SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }).then((res: Response) => {
                return res.json();
            });
    }
}

