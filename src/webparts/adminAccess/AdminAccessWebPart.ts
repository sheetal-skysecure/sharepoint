import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    type IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import AdminPortal from './components/AdminPortal';
import * as strings from 'AdminAccessWebPartStrings';
import { SharePointService } from '../learningCenter/services/SharePointService';

export interface IAdminAccessWebPartProps {
    description: string;
}

export default class AdminAccessWebPart extends BaseClientSideWebPart<IAdminAccessWebPartProps> {
    private _adminAccessState: { canAccessAdmin: boolean; isOwner: boolean; userRole: string } = {
        canAccessAdmin: false,
        isOwner: false,
        userRole: 'Unknown'
    };


    protected onInit(): Promise<void> {
        return super.onInit().then(async () => {
            const siteUrl = this.context.pageContext.web.absoluteUrl;
            SharePointService.init(siteUrl, this.context.spHttpClient, this.context);

            try {
                const accessState = await SharePointService.getCurrentUserAdminAccess(this.context.pageContext.user.email, true);
                const resolvedRole = accessState.currentUserRole === 'Unknown'
                    ? 'Visitor'
                    : accessState.currentUserRole;

                this._adminAccessState = {
                    canAccessAdmin: accessState.canAccessAdmin,
                    isOwner: resolvedRole === 'Owner',
                    userRole: resolvedRole
                };
            } catch (error) {
                console.error('[AdminAccessWebPart] Failed to determine admin access', error);
                this._adminAccessState = {
                    canAccessAdmin: false,
                    isOwner: false,
                    userRole: 'Visitor'
                };
            }
        });
    }

    public render(): void {
        if (!this._adminAccessState.canAccessAdmin) {
            const element: React.ReactElement = React.createElement(
                'div',
                {
                    style: {
                        minHeight: '70vh',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        background: '#f8fafc',
                        padding: '2rem'
                    }
                },
                React.createElement(
                    'div',
                    {
                        style: {
                            maxWidth: '520px',
                            width: '100%',
                            background: '#ffffff',
                            border: '1px solid #e2e8f0',
                            borderRadius: '24px',
                            boxShadow: '0 20px 40px -20px rgba(15, 23, 42, 0.25)',
                            padding: '2rem',
                            textAlign: 'center'
                        }
                    },
                    React.createElement('h2', { style: { margin: '0 0 1rem', color: '#0f172a' } }, 'Access Restricted'),
                    React.createElement(
                        'p',
                        { style: { margin: '0 0 1.5rem', color: '#475569', lineHeight: 1.6 } },
                        'You do not have permission to access the Admin Portal. Only SharePoint Owners and Members are allowed.'
                    ),
                    React.createElement(
                        'button',
                        {
                            type: 'button',
                            onClick: () => {
                                window.location.assign(this.context.pageContext.web.absoluteUrl);
                            },
                            style: {
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                padding: '0.85rem 1.25rem',
                                background: 'linear-gradient(135deg, #0ea5e9 0%, #10b981 100%)',
                                color: '#ffffff',
                                borderRadius: '14px',
                                fontWeight: 800,
                                textDecoration: 'none',
                                border: 'none',
                                cursor: 'pointer'
                            }
                        },
                        'Go to Home'
                    )
                )
            );

            ReactDom.render(element, this.domElement);
            return;
        }

        const element: React.ReactElement = React.createElement(
            AdminPortal,
            {
                userDisplayName: this.context.pageContext.user.displayName,
                userEmail: this.context.pageContext.user.email,
                isOwner: this._adminAccessState.isOwner,
                canAccessAdmin: this._adminAccessState.canAccessAdmin,
                userRole: this._adminAccessState.userRole,
                context: this.context
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        const { semanticColors } = currentTheme;

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
