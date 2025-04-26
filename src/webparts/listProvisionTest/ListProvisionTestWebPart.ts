import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, Icon } from '@fluentui/react';

import * as strings from 'ListProvisionTestWebPartStrings';
import ListProvisionTest from './components/ListProvisionTest';

interface IListItem {
  ID: string;
  Title: string;
  Status: string;
  AssignedTo: {
    Title: string;
  };
}

export interface IListProvisionTestWebPartProps {
  projectsListExists: boolean;
  userDisplayName?: string;
}

export default class ListProvisionTestWebPart extends BaseClientSideWebPart<IListProvisionTestWebPartProps> {
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private isDialogVisible: boolean = false;
  private items: IListItem[] = [];

  protected async onInit(): Promise<void> {
    await super.onInit();
    await Promise.all([
      this.getCurrentUserInfo(),
      this.checkProjectsList()
    ]);
    
    if (this.properties.projectsListExists) {
      await this.fetchListItems();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private async getCurrentUserInfo(): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const userProperties = await response.json();
        this.properties.userDisplayName = userProperties.DisplayName || userProperties.PreferredName;
      }
    } catch (error) {
      console.error('Error fetching user info:', error);
      this.properties.userDisplayName = 'User';
    }
  }

  private async checkProjectsList(): Promise<void> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')`,
        SPHttpClient.configurations.v1
      );

      this.properties.projectsListExists = response.ok;
    } catch {
      this.properties.projectsListExists = false;
    }
  }

  private getRandomStatus(): string {
    const statuses = ['Not Started', 'In Progress', 'Completed'];
    return statuses[Math.floor(Math.random() * statuses.length)];
  }

  private getRandomTitle(): string {
    const prefixes = ['Project', 'Initiative', 'Task', 'Development'];
    const descriptors = ['Alpha', 'Beta', 'Phase 1', 'Planning', 'Implementation'];
    const areas = ['UI', 'Backend', 'Database', 'Testing', 'Documentation'];
    
    const prefix = prefixes[Math.floor(Math.random() * prefixes.length)];
    const descriptor = descriptors[Math.floor(Math.random() * descriptors.length)];
    const area = areas[Math.floor(Math.random() * areas.length)];
    
    return `${prefix} ${descriptor} - ${area}`;
  }

  private async createSampleItems(): Promise<void> {
    try {
      // Get current user info to use as AssignedTo
      const userResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );
      
      if (!userResponse.ok) {
        throw new Error('Failed to get current user');
      }

      const currentUser = await userResponse.json();
      
      // Create 5 sample items
      for (let i = 0; i < 5; i++) {
        const response = await this.context.spHttpClient.post(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            },
            body: JSON.stringify({
              Title: this.getRandomTitle(),
              Status: this.getRandomStatus(),
              AssignedToId: currentUser.Id
            })
          }
        );

        if (!response.ok) {
          console.error(`Failed to create sample item ${i + 1}`);
        }
      }
    } catch (error) {
      console.error('Error creating sample items:', error);
    }
  }

  private async provisionList(): Promise<void> {
    try {
      const createListResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify({
            Title: "ProjectsNew",
            BaseTemplate: 100,
            AllowContentTypes: true,
            ContentTypesEnabled: true
          })
        }
      );

      if (!createListResponse.ok) {
        throw new Error(`Failed to create list: ${createListResponse.statusText}`);
      }

      await new Promise(resolve => setTimeout(resolve, 1000));

      const statusFieldResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '3.0'
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.FieldChoice' },
            'Title': 'Status',
            'FieldTypeKind': 6,
            'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, 'results': ['Not Started', 'In Progress', 'Completed'] },
            'DefaultValue': 'Not Started'
          })
        }
      );

      if (!statusFieldResponse.ok) {
        throw new Error('Failed to create Status field');
      }

      const assignedToFieldResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '3.0'
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.FieldUser' },
            'Title': 'AssignedTo',
            'FieldTypeKind': 20,
            'AllowMultipleValues': false
          })
        }
      );

      if (!assignedToFieldResponse.ok) {
        throw new Error('Failed to create AssignedTo field');
      }

      // Get the default view
      const viewResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/views/getbytitle('All Items')`,
        SPHttpClient.configurations.v1
      );

      if (!viewResponse.ok) {
        throw new Error('Failed to get default view');
      }

      // Update the default view to include our fields
      await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/views/getbytitle('All Items')/viewfields/addviewfield('Status')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '3.0',
            'X-HTTP-Method': 'POST'
          }
        }
      );

      await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/views/getbytitle('All Items')/viewfields/addviewfield('AssignedTo')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '3.0',
            'X-HTTP-Method': 'POST'
          }
        }
      );

      // Create sample items
      await this.createSampleItems();

      this.properties.projectsListExists = true;
      this.context.propertyPane.refresh();
      alert("ProjectsNew list has been created successfully with sample items!");

    } catch (error) {
      console.error('Error creating list:', error);
      alert("Failed to create ProjectsNew list. Check console for details.");
    }
  }

  private showDialog(): void {
    this.isDialogVisible = true;
    this.render();
  }

  private hideDialog(): void {
    this.isDialogVisible = false;
    this.render();
  }

  private async fetchListItems(): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectsNew')/items?$select=ID,Title,Status,AssignedTo/Title&$expand=AssignedTo`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        this.items = data.value;
        this.render();
      }
    } catch (error) {
      console.error('Error fetching list items:', error);
    }
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(
      'div',
      {},
      [
        React.createElement(
          ListProvisionTest,
          {
            projectsListExists: this.properties.projectsListExists,
            userDisplayName: this.properties.userDisplayName || 'User',
            onConfigureClick: () => {
              this.context.propertyPane.open();
            },
            items: this.items
          }
        ),
        this.isDialogVisible && React.createElement(
          Dialog,
          {
            hidden: false,
            onDismiss: () => this.hideDialog(),
            dialogContentProps: {
              type: DialogType.normal,
              title: '',
              showCloseButton: true,
              styles: {
                content: {
                  width: '800px'
                }
              }
            },
            modalProps: {
              isBlocking: true,
              styles: { 
                main: { 
                  maxWidth: '800px !important',
                  minWidth: '800px !important',
                  minHeight: '300px'
                } 
              }
            }
          },
          [
            React.createElement(
              'div',
              { style: { textAlign: 'center', marginBottom: '20px' } },
              React.createElement(Icon, {
                iconName: "Warning",
                styles: {
                  root: {
                    fontSize: '64px',
                    color: '#f0ad4e',
                    marginBottom: '20px'
                  }
                }
              })
            ),
            React.createElement(
              'h3',
              { style: { textAlign: 'center', margin: '0 0 20px 0', fontSize: '16px' } },
              [
                'This action will create a new list in ',
                React.createElement(
                  'span',
                  { style: { color: '#0078d4' } },
                  this.context.pageContext.web.absoluteUrl
                ),
                '.'
              ]
            ),
            React.createElement(
              'h3',
              { style: { textAlign: 'center', margin: '0 0 20px 0', fontSize: '16px' } },
              'Ok to proceed?'
            ),
            React.createElement(
              DialogFooter,
              {},
              [
                React.createElement(PrimaryButton, {
                  onClick: async () => {
                    this.hideDialog();
                    await this.provisionList();
                  },
                  text: "OK"
                }),
                React.createElement(DefaultButton, {
                  onClick: () => this.hideDialog(),
                  text: "Cancel"
                })
              ]
            )
          ]
        )
      ]
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const provisionButton = !this.properties.projectsListExists ? 
      PropertyPaneButton('provisionList', {
        text: "Build the list for me",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: () => this.showDialog()
      }) : null;

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
                ...(!this.properties.projectsListExists ? [
                  PropertyPaneLabel('projectsListStatus', {
                    text: "A required list is missing."
                  })
                ] : []),
                ...(provisionButton ? [provisionButton] : [])
              ]
            }
          ]
        }
      ]
    };
  }
}


