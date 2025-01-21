import { Version } from '@microsoft/sp-core-library';
//import * as React from 'react';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CruddemoWebPart.module.scss';
import * as strings from 'CruddemoWebPartStrings';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ICruddemoWebPartProps {
  description: string;
}

export default class CruddemoWebPart extends BaseClientSideWebPart<ICruddemoWebPartProps> {

  private locations: string[] = [];
  private isLocationsLoaded: boolean = false;  // Flag to track when locations are fully loaded

  public onInit(): Promise<void> {
    // Initialize locations fetch once when the web part is loaded
    return super.onInit().then(_ => {
      this._getLocations(); // Fetch Locations from SharePoint list when the web part is initialized
    });
  }

  public render(): void {
    // Prevent rendering until locations are loaded
    if (!this.isLocationsLoaded) {
      this.domElement.innerHTML = `<section class="${styles.cruddemo} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
                                      <div>Loading Locations...</div>
                                    </section>`;
      return; // Return early, don't render the form yet
    }

    // Now render the form with populated locations
    this.domElement.innerHTML = `
      <section class="${styles.cruddemo} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>

        <div>
          <table border='5' bgcolor='aqua'>
            <tr>
              <td>Please Enter Software ID </td>
              <td><input type='text' id='txtID' />
              <td><input type='submit' id='btnRead' value='Read Details' />
              </td>
            </tr>

            <tr>
              <td>Software Title</td>
              <td><input type='text' id='txtSoftwareTitle' /></td>
            </tr>

            <tr>
              <td>Software Name</td>
              <td><input type='text' id='txtSoftwareName' /></td>
            </tr>

            <tr>
              <td>Location</td>
              <td>
                <select id="ddLocation">
                  ${this.locations.map(location => `<option value="${location}">${location}</option>`).join('')}
                </select>
              </td>
            </tr>
             <tr>
              <td>Job Title</td>s
              <td>
                <input type='text' id='txtJobTitle' value="Front End SharePoint Developer" disabled /></td>
                </td>
            </tr>
             <tr>
              <td>Name of operator</td>
              <td>
                <input type='text' id='txtOperator' value="${escape(this.context.pageContext.user.displayName)}" disabled /></td>
                </td>
            </tr>
             <tr>
              <td>Possible?</td>
              <td>
                  <select id="ddPossible">
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                  </select>
                </td>
            </tr>

            <tr>
              <td>Software Version</td>
              <td><input type='text' id='txtSoftwareVersion' /></td>
            </tr>

            <tr>
              <td>Software Description</td>
              <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea> </td>
            </tr>

            <tr>
              <td colspan='2' align='center'>
                <input type='submit' value='Insert Item' id='btnSubmit' />
              </td>
            </tr>
          </table>
        </div>
        <div id="divStatus"/>
      </section>`;

    // Bind events after the form is rendered
    this._bindEvents();
  }

  // Fetch Locations from the Locations list
  private _getLocations(): void {
    if (this.isLocationsLoaded) {
      return; // If locations are already loaded, don't fetch again
    }

    const siteUrl = this.context.pageContext.site.absoluteUrl;
     const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Locations')/items?$top=1000`;
    //const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Locations')/items?$top`;

    this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data) => {
        // Update locations array with the fetched data
        this.locations = data.value.map((item: { Title: string }) => item.Title);
        this.isLocationsLoaded = true;  // Mark locations as loaded
        this.render();  // Re-render the component after locations are loaded
      })
      .catch((error) => {
        console.error('Error fetching Locations: ', error);
        this.isLocationsLoaded = false;
        this.render();  // Re-render with error state
      });
  }

  // Bind the Events
  private _bindEvents(): void {
    // Add a new item
    this.domElement.querySelector('#btnSubmit')!.addEventListener('click', () => { this.addListItem(); });
  }

  // Add the list item to the SoftwareCatalog list
  private addListItem(): void {
    var softwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;
    var softwareName = (document.getElementById("txtSoftwareName") as HTMLInputElement).value;
    var softwareVersion = (document.getElementById("txtSoftwareVersion") as HTMLInputElement).value;
    var softwareDescription = (document.getElementById("txtSoftwareDescription") as HTMLInputElement).value;
    var possible = (document.getElementById("ddPossible") as HTMLInputElement).value;
    var jobtitle = (document.getElementById("txtJobTitle") as HTMLInputElement).value;
    var possiblevalue: boolean = possible === 'Yes';
    const location = (document.getElementById("ddLocation") as HTMLSelectElement).value;

    const operator = (document.getElementById("txtOperator") as HTMLInputElement).value;

    // Get the Location ID from Locations List based on the selected Title
    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    // We need to get the ID of the selected Location
    const locationId = this.locations.indexOf(location) + 1; // Location ID corresponds to the array index

    
    const itemBody: any = {
      "Title": softwareTitle,
      "SoftwareName": softwareName,
      "SoftwareVersion": softwareVersion,
      "SoftwareDescription": softwareDescription,
      "Possible": possiblevalue,
      "JobTitle": jobtitle,
      "LocationId": locationId, // Reference the Location ID here,
      "Operators": operator,
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    // RESTful API Call to insert a new item in the SoftwareCatalog list
    this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          let statusmessage = this.domElement.querySelector('#divStatus') as Element;
          statusmessage.innerHTML = "List Item has been created successfully.";
          this.clear();
        } else {
          let statusmessage = this.domElement.querySelector('#divStatus') as Element;
          statusmessage.innerHTML = "An error has occurred: " + response.status + " - " + response.statusText;
          response.text().then(errorText => {
            console.error("Error Details:", errorText); // Log detailed error message
          });
        }
      });
  }

  // Clear the input fields
  private clear(): void {
    (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value = '';
    (document.getElementById("txtSoftwareName") as HTMLInputElement).value = '';
    (document.getElementById("txtSoftwareVersion") as HTMLInputElement).value = '';
    (document.getElementById("txtSoftwareDescription") as HTMLInputElement).value = '';
    (document.getElementById("ddLocation") as HTMLSelectElement).value = '';
    
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
