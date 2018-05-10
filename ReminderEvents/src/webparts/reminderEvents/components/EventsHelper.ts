import * as moment from 'moment';
declare var $;
export default class EventsHelper {
    /*Common Methods - Start*/

    public getParameterByName(name, url) {
        if (!url) url = window.location.href;
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

    public formatDate(date) {
        var DateObj = new Date(date);
        var arr = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
        var dd: any = DateObj.getDate();
        var mm = DateObj.getMonth();
        var yyyy = DateObj.getFullYear();
        if (dd < 10) { dd = `0${dd}`; }
        return arr[mm] + ' ' + (dd - 1) + ', ' + yyyy;
    }

    public formatAMPM(date) {
        let DateObj: any = new Date(Date.parse(date));
        var hours = DateObj.getHours();
        var minutes = DateObj.getMinutes();
        var ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12; // the hour '0' should be '12'
        hours = hours < 10 ? '0' + hours : hours;
        minutes = minutes < 10 ? '0' + minutes : minutes;
        var strTime = hours + ':' + minutes + ' ' + ampm;
        //strTime=this.removeSpecialChars(strTime).replace("00", ":00").replace("30", ":30");
        return strTime;
    }

    public GetTrimmedText(TitleText: string, Size: number) {
        var TrimmedText = TitleText;
        if (TitleText == null)
            return TitleText;
        if (TitleText.length > Size) {
            TrimmedText = this.removeHtmlAndTrimStringWithEllipsis(TitleText, Size + 3).replace("...", "");
        }
        if (TitleText.length > Size && TitleText.substr(Size, 1) != " " && TrimmedText.lastIndexOf(" ") > 0) {
            TrimmedText = TrimmedText.substr(0, TrimmedText.lastIndexOf(" ")) + " ...";
        } else if (TitleText.length > Size && TitleText.substr(Size, 1) == " ") {
            TrimmedText = TrimmedText + " ...";
        }
        return TrimmedText;
    }


    public removeHtmlAndTrimStringWithEllipsis(d, b) {
        var a = "";
        a = this.removeHtml(d);
        if (b >= 3 && !this.isNullOrUndefinedOrEmpty(a))
            if (a.length > b) {
                var c = "...";
                a = a.substr(0, Math.max(b - c.length, 0));
                a += c;
            }
        return a;
    }

    public removeHtml(b) {
        var a = document.createElement("div");
        a.innerHTML = b;
        this.removeStyleChildren(a);
        return this.GetInnerText(a);
    }

    public GetInnerText(a) {
        return typeof a.textContent !== "undefined" && a.textContent !== null ? a.textContent : typeof a.innerText !== "undefined" ? a.innerText : undefined;
    }

    public removeStyleChildren(d) {
        var b = d.getElementsByTagName("style");
        if (!this.isNullOrUndefinedOrEmpty(b))
            for (var c = b.length - 1; c >= 0; c--) {
                var a = b[c];
                return !this.isNullOrUndefinedOrEmpty(a) && !this.isNullOrUndefinedOrEmpty(a.parentNode) && a.parentNode.removeChild(a);
            }
    }



    public getUTCtoLocalTime(date) {
        var utc_date = new Date(date);
        var dt = new Date(utc_date.toUTCString()).toLocaleString([], { hour12: true }).toString().replace(',', '');
        if (dt) {
            var datestr = dt.split(' ')[0];
            var hh = dt.split(' ')[1].split(':')[0];
            var mm = dt.split(' ')[1].split(':')[1];
            var meridiem = dt.split(' ')[2];
            dt = hh + ':' + mm + ' ' + meridiem;
        }
        return dt;
    }

    public isEmptyString(value: string): boolean {
        return value === null || typeof value === "undefined" || !value.length;
    }


    public getTermString(owstaxIdHashTags) {
        if (owstaxIdHashTags == undefined || owstaxIdHashTags == null || owstaxIdHashTags == "") {
            return "";
        }
        owstaxIdHashTags = owstaxIdHashTags.split("\n\n").join("^^");
        var Terms = owstaxIdHashTags.split("|");
        var checkIndex = 3;
        var sTerms = "";
        for (var i = 1; i < Terms.length; i++) {
            if (Terms[i].indexOf("#") == -1) {
                if (Terms[i].indexOf("^^") != -1) {
                    sTerms += Terms[i].substr(0, Terms[i].indexOf("^^")) + ";";
                }
                else {
                    sTerms += Terms[i];
                }
            }
        }
        if (sTerms.charAt(sTerms.length - 1) == ";") {
            sTerms = sTerms.substr(0, sTerms.length - 1);
        }
        if (sTerms.indexOf(";GTSet") != -1)
            sTerms = sTerms.replace(";GTSet", "");
        sTerms = sTerms.replace(/;/g, " | ");
        return sTerms;
    }


    public isNullOrUndefinedOrEmpty(b) {
        var c = null, a = b;
        return a === c || typeof a === "undefined" || b === "";
    }

    public loadScript(url, callback) {

        var script = document.createElement("script");
        script.type = "text/javascript";

        script.onload = () => {
            callback();
        };

        script.src = url;
        document.getElementsByTagName("head")[0].appendChild(script);
    }

    public ValidateProperty(value: string) {
        if (!(value === null || typeof value === "undefined" || !value.length)) {
            if (value.toLocaleLowerCase().indexOf('<script>') == -1 && value.toLocaleLowerCase().indexOf('<img') == -1) {
                if (value.length > 35) {
                    return 'Should be less than 35 characters';
                }
                else
                    return '';
            }
            else {
                return 'Script,img tags not allowed';
            }
        }
    }

    public loadCSS(url) {
        var loaded = false;
        var ss = document.styleSheets;
        for (var i = 0, max = ss.length; i < max; i++) {
            if (ss[i].href != null && ss[i].href.toLowerCase() == url.toLowerCase()) {
                loaded = true;
                return;
            }
        }
        if (!loaded) {
            var link = document.createElement("link");
            link.rel = "stylesheet";
            link.href = url;
            document.getElementsByTagName("head")[0].appendChild(link);
        }
    }

    public ConvertUTCToLocalDateTime(siteURL, UTCDateStr) {
        var dtFormat = "";
        var that = this;
        if (this.isNullOrUndefinedOrEmpty(UTCDateStr))
            return "";
        var momentDate = moment.utc(UTCDateStr);
        console.log("MomentDate", momentDate);
        var month_names_short = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var DateObj = new Date(UTCDateStr);
        var utcDate = DateObj.getDate();
        var utcMonth = DateObj.getMonth();
        var utcYear = DateObj.getFullYear();
        var localString = month_names_short[utcMonth] + ' ' + utcDate + ', ' + utcYear;
        console.log("LocalDate", localString);
        $.ajax({
            url: siteURL +
            "/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@date)?@date='" +
            UTCDateStr +
            "'",
            async: false,
            method: "GET",
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: (data) => {
                var regionalDate = data.d.UTCToLocalTime;
                if (regionalDate.toUpperCase().indexOf('Z') === -1) {
                    regionalDate = regionalDate + 'Z';
                }
                var localdate = new Date(regionalDate);
                var mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                var locDate = localdate.getUTCDate();
                var locMonth = localdate.getUTCMonth();
                var locYear = localdate.getUTCFullYear();
                var hours = localdate.getUTCHours();
                var meridiem = 'AM';
                if (hours == 0) { //At 00 hours we need to show 12 am
                    hours = 12;
                }
                else if (hours > 12) {
                    hours = hours % 12;
                    meridiem = 'PM';
                }
                var minutes: any = localdate.getUTCMinutes();
                if (minutes < 10)
                    minutes = '0' + minutes;
                var hoursFormat = hours + ':' + minutes + ' ' + meridiem;
                dtFormat = that.removeSpecialChars(hoursFormat).replace("00", ":00").replace("30", ":30");
                dtFormat = mon[locMonth] + ' ' + locDate + ', ' + locYear + "|" + hoursFormat;

            },
            error: (errMsg) => {
                if (errMsg.responseText) {
                    console.log(errMsg.responseText);
                }
            }
        });
        return dtFormat;
    }

    public removeSpecialChars(str) {
        return str.replace(/(?!\w|\s)./g, '')
            .replace(/\s+/g, ' ')
            .replace(/^(\s*)([\W\w]*)(\b\s*$)/g, '$2');
    }
}

