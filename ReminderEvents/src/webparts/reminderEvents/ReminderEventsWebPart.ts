import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import ReminderEvents from './components/ReminderEvents';
import { IReminderEventsProps } from './components/IReminderEventsProps';
import helper from "./components/EventsHelper";
declare var $;

export interface IReminderEventsWebPartProps {
    description: string;
}

export default class ReminderEventsWebPart extends BaseClientSideWebPart<IReminderEventsWebPartProps> {
    private _cdnUrl: string;
    private _helper: helper;
    public constructor() {
        super();
        this._helper = new helper();
    }
    protected onInit(): Promise<void> {
        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Loading..");
        this._cdnUrl = this.context.manifest.loaderConfig.internalModuleBaseUrls ? this.context.manifest.loaderConfig.internalModuleBaseUrls[0].split("#")[0].toLocaleLowerCase().replace("/js/dist/ReminderEvents", "").replace(/\/$/, '') : "";
        return super.onInit();
    }
    public render(): void {
        const element: React.ReactElement<IReminderEventsProps> = React.createElement(
            ReminderEvents,
            {
                context: this.context,
                WPProperties: this.properties,
                InstanceId: this.context.instanceId,
                CDNURL: this._cdnUrl
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: [
                                PropertyPaneTextField('WebPartTitle', {
                                    label: "Title",
                                    onGetErrorMessage: this._helper.ValidateProperty.bind(this)
                                }),
                                PropertyPaneDropdown('IsBoxed', {
                                    label: "Box",
                                    options: [
                                        { key: 'Yes', text: 'Yes' },
                                        { key: 'No', text: 'No' }
                                    ],
                                    selectedKey: 'No'
                                }),
                                PropertyPaneDropdown('NumberOfItems', {
                                    label: "Number Of Items",
                                    options: [
                                        { key: '1', text: '1' },
                                        { key: '2', text: '2' },
                                        { key: '3', text: '3' },
                                        { key: '4', text: '4' },
                                        { key: '5', text: '5' },
                                        { key: '6', text: '6' },
                                        { key: '7', text: '7' },
                                        { key: '8', text: '8' },
                                        { key: '9', text: '9' },
                                        { key: '10', text: '10' },
                                        { key: '11', text: '11' },
                                        { key: '12', text: '12' }
                                    ],
                                    selectedKey: '4'
                                }),
                                PropertyPaneDropdown('EventType', {
                                    label: "Event Type",
                                    options: [
                                        { key: 'All', text: 'All' },
                                        { key: 'Reminder', text: 'Reminder' }
                                    ],
                                    selectedKey: 'Reminder'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}