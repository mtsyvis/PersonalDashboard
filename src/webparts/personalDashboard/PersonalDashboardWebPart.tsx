import * as React from 'react';
import * as ReactDom from "react-dom";
import { Provider } from "react-redux";
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PersonalDashboardWebPartStrings';
import PersonalDashboard from './components/PersonalDashboard';

import { store } from "./store/store";

export interface IPersonalDashboardWebPartProps {
    description: string;
}

export default class PersonalDashboardWebPart extends BaseClientSideWebPart<IPersonalDashboardWebPartProps> {

    private _isDarkTheme: boolean = false;

    public render(): void {
        ReactDom.render(
            <Provider store={store}>
                <PersonalDashboard
                    serviceScope={this.context.serviceScope}
                    hasTeamsContext={!!this.context.sdks.microsoftTeams}
                    userDisplayName={this.context.pageContext.user.displayName} />
            </Provider>,
            this.domElement
        );
    }

    protected onInit(): Promise<void> {
        return super.onInit();
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        this._isDarkTheme = !!currentTheme.isInverted;
        const {
            semanticColors
        } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }

    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
