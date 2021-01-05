import * as React from "react";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { TaskModuleDimension } from "@microsoft/teams-js";
import WelcomeCard from "../../acPrototypeBot/dialogs/WelcomeDialog";

/**
 * State for the acPrototypeTabTab React component
 */
export interface IAcPrototypeTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the acPrototypeTabTab React component
 */
export interface IAcPrototypeTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the acPrototype Tab content page
 */
export class AcPrototypeTab extends TeamsBaseComponent<IAcPrototypeTabProps, IAcPrototypeTabState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
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
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };

        const onClick = () => {
            microsoftTeams.tasks.startTask({
                height: TaskModuleDimension.Medium,
                width: TaskModuleDimension.Medium,
                card: WelcomeCard,
                title: "test",
            });
        };

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>This is your tab from index.html</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                {this.state.entityId}
                            </div>
                            <div style={styles.section}>
                                <PrimaryButton onClick={onClick}>Task Module Button</PrimaryButton>
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>
                                (C) Copyright proto
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
}
