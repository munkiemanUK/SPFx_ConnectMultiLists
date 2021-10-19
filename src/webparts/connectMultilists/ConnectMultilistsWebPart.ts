import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {IPropertyPaneConfiguration,PropertyPaneTextField} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ConnectMultilistsWebPart.module.scss';
import * as strings from 'ConnectMultilistsWebPartStrings';

import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment,EnvironmentType} from '@microsoft/sp-core-library';
import { ComboBox, IComboBoxOption, IComboBox, PrimaryButton } from 'office-ui-fabric-react/lib/index';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Web } from "@pnp/sp/presets/all";
import { getGUID } from "@pnp/common";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// import node module external libraries
import * as $ from 'jquery';
import * as React from 'react';
require('bootstrap');
require('./styles/custom.css');

// setup global variables
let assessmentType: string='';

export interface IConnectMultilistsWebPartProps {
  description: string;
}

export interface AuditLists {
  value: AuditQuestionsList[];
}

export interface AuditQuestionsList {
  Id: string;
  Section: string;
  Assessment: string;
  Question_Number: string;
  Title: string;
  Min_Outcome: string;
}

export default class ConnectMultilistsWebPart extends BaseClientSideWebPart<IConnectMultilistsWebPartProps> {

  private async _getConsultationData(): Promise<AuditLists> {
    const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('Audit%20Tool%20Questions')/Items?$filter=Section eq 'Consultation Records'`, SPHttpClient.configurations.v1);
    return await response.json();
  }

  private _renderList(items: AuditQuestionsList[]): void {
    let html: string = '';
    $('#assessment a').on('click', function(){
      alert($(this).text());
      assessmentType=$(this).text();
    });
    
    items.forEach((item: AuditQuestionsList) => {       
      let assessment: string=item.Assessment;
      console.log(assessment+" "+assessmentType);
           
        html += `
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span class="ms-font-l">${item.Section}</span>
            <span class="ms-font-l">${item.Assessment}</span>
            <span class="ms-font-l">${item.Question_Number}</span>
            <span class="ms-font-l">${item.Title}</span>
            <span class="ms-font-l">${item.Min_Outcome}</span>
          </li>
        </ul>`;
    
    });
  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    // Local environment
    //if (Environment.type === EnvironmentType.Local) {
    //  this._getMockListData().then((response) => {
    //    this._renderList(response.value);
    //  });
    //}
    //else 
    //$('#assessment a').on('click', function(){
    //  alert($(this).text());
    //  assessmentType=$(this).text();
    //});

    if (Environment.type == EnvironmentType.SharePoint ||
             Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getConsultationData()
        .then((response) => {
          alert(response.value);
          this._renderList(response.value);
        });
    }
  }  

  //public onChangeSelect(event: any): void {
  //  this.setState({ ItemCountry: event.target.value });
  //}

  public render(): void {
    let bootstrapCssURL = "https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css";
    let fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);
    
    this.domElement.innerHTML = `
      <div class="${ styles.connectMultilists }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using web parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              
              <div class="dropdown mr-1">
                <button type="button" class="btn dropdown-toggle dropdown-toggle-split text-white" style="background-color: #545487;" id="dropdownAssessment" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" data-reference="parent">
                    Assessment
                </button>
                <div id="assessmentSelected"></div>
                <div id="assessment" class="dropdown-menu" aria-labelledby="dropdownAssessment">
                    <a class="dropdown-item" href="#">Audiometry</a>
                    <a class="dropdown-item" href="#">Fitness for Task</a>
                    <a class="dropdown-item" href="#">HAVs Tier 1 & 2</a>
                    <a class="dropdown-item" href="#">HAVs Tier 3</a>
                    <a class="dropdown-item" href="#">HAVs Tier 4</a>
                    <a class="dropdown-item" href="#">Immunisation</a>
                    <a class="dropdown-item" href="#">Management Referral</a>
                    <a class="dropdown-item" href="#">Safety Critical</a>
                    <a class="dropdown-item" href="#">Skin</a>
                </div>
                <div className="col-3"> 
              </div>                
              </div>
              <p class="${ styles.description }">Loading from ${escape(this.context.pageContext.web.title)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
          <div id="spListContainer" />
        </div>
      </div>`;
  
    this._renderListAsync();
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
