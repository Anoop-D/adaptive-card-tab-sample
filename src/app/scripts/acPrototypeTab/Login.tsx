import * as React from "react";
import {
    TeamsThemeContext,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the defaultTabTab React component
 */
export interface ILoginState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the defaultTabTab React component
 */
export interface ILoginProps extends ITeamsBaseComponentProps {

}
/**
 * Implementation of the DefaultTab content page
 */
export class Login extends TeamsBaseComponent<ILoginProps, ILoginState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize(),
        });

        const state = Math.random().toString(); // _guid() is a helper function in the sample

        // Go to the Azure AD authorization endpoint
        const queryParams = {
            client_id: "eda13c8b-ec36-4ef5-a600-999c9531a536",
            response_type: "id_token token",
            response_mode: "fragment",
            resource: "https://graph.microsoft.com",
            redirect_uri: "https://andhillo-relay.servicebus.windows.net:443/MININT-S5EDEDH/acPrototypeTab/redirect.html",
            nonce: Math.random().toString(),
            state,
        };

        const authorizeEndpoint = "https://login.microsoftonline.com/" + "72f988bf-86f1-41af-91ab-2d7cd011db47" + "/oauth2/authorize?" + this.toQueryString(queryParams);
        window.location.assign(authorizeEndpoint);
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        return (
            <TeamsThemeContext.Provider value={context}>
            </TeamsThemeContext.Provider>
        );
    }

    // Build query string from map of query parameter
    private toQueryString(queryParams) {
        // tslint:disable-next-line: prefer-const
        let encodedQueryParams: string[] = [];
        for (const key of Object.keys(queryParams)) {
            encodedQueryParams.push(key + "=" + encodeURIComponent(queryParams[key]));
        }
        return encodedQueryParams.join("&");
    }
}
