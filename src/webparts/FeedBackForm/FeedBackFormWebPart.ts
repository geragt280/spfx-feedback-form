import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import * as strings from "FeedBackFormWebPartStrings";
import FeedBackForm from "./components/FeedBackForm";
import { IFeedBackFormProps } from "./components/IFeedBackFormProps";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import {
  IColumnReturnProperty,
  PropertyFieldColumnPicker,
  PropertyFieldColumnPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker";

import { SPFI, spfi, SPFx } from "@pnp/sp";

export interface IFeedBackFormWebPartProps {
  Title: string;
  colParticipant: string;
  colComments: string;
  colSupportType: string;
  FeedbacklistID: string;
  colTitle: string;
  headingBackColor: string;
  enableReEnterFormLink: boolean;
}

export default class FeedBackFormWebPart extends BaseClientSideWebPart<IFeedBackFormWebPartProps> {
 private spList:SPFI;
  public render(): void {
    const element: React.ReactElement<IFeedBackFormProps> = React.createElement(
      FeedBackForm,
      {
        title: this.properties.Title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.Title = value          
        },
        sp:this.spList,
        SupportType:this.properties.colSupportType,
        context:this.context,
        FeedbackListID:this.properties.FeedbacklistID,
        colComments: this.properties.colComments,
        colSupportType: this.properties.colSupportType,
        colTitle: this.properties.colTitle,
        headingBackColor: this.properties.headingBackColor,
        enableReEnterFormLink: this.properties.enableReEnterFormLink
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    
    return this._getEnvironmentMessage().then((message) => {
       this.spList = spfi().using(SPFx(this.context));
      //this._environmentMessage = message;
    });
    
    
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    //this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.description,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("Title", {
                  //label: strings.DescriptionFieldLabel,
                  label: strings.Title,
                }),
                PropertyFieldListPicker("FeedbacklistID", {
                  label: strings.lbldrpFeedbackList,
                  selectedList: this.properties.FeedbacklistID,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                  baseTemplate: 100,
                }),
                PropertyFieldColumnPicker("colComments", {
                  label: strings.lblDrpCommentsColumn,
                  context: this.context,
                  selectedColumn: this.properties.colComments,
                  listId: this.properties.FeedbacklistID,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  filter: "Hidden eq false and ReadOnlyField eq false and FieldTypeKind eq 3",
                  key: "colComments",
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"],
                }),
                PropertyFieldColumnPicker("colSupportType", {
                  label: strings.lblDrpFeedbackTypeColumn,
                  context: this.context,
                  selectedColumn: this.properties.colSupportType,
                  listId: this.properties.FeedbacklistID,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  filter: "Hidden eq false and ReadOnlyField eq false and FieldTypeKind eq 6",
                  key: "colSupportType",
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"],
                }),
                PropertyFieldColorPicker('headingBackColor', {
                  label: 'Header color',
                  selectedColor: this.properties.headingBackColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'headingBackColorFieldId'
                }),
                PropertyPaneToggle("enableReEnterFormLink",{
                  label: strings.enableReEnterFormLink,
                  checked: this.properties.enableReEnterFormLink
                })
              ],
            },
          ],
        },
      ],
    };
  }
}
