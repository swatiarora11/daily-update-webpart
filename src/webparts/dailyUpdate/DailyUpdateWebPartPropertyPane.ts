import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

import * as strings from "DailyUpdateWebPartStrings";

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

import { IDailyUpdateWebPartProps } from "./DailyUpdateWebPart";
import { PropertyFieldSitePicker } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Utils } from "../../services";
import { IGroup } from "../../models";

export class DailyUpdateWebPartPropertyPane {
  private _groupsfields: IPropertyPaneField<any>[] = [];
  private context = undefined;

  private onPropertyPaneFieldChanged: (propertyPath: string, oldValue: any, newValue: any) => Promise<void>;
  private properties: IDailyUpdateWebPartProps;

  private groups = undefined;

  constructor(
    context: WebPartContext,
    properties: IDailyUpdateWebPartProps,
    groups: IGroup[],
    onPropertyPaneFieldChanged: (propertyPath: string, oldValue: any, newValue: any) => Promise<void>
  ) {
    this.context = context;
    this.properties = properties;
    this.groups = groups;

    this.onPropertyPaneFieldChanged = onPropertyPaneFieldChanged;
  }

  private _getGroupFields = async () => {
    this._groupsfields = [
      PropertyPaneTextField("title", {
        label: strings.TitleFieldLabel,
      }),
      PropertyPaneTextField("iconProperty", {
        label: strings.IconPropertyFieldLabel,
      }),
      PropertyPaneTextField("adminUpn", {
        label: strings.AdminUpnPropertyFieldLabel,
      }),
      PropertyPaneDropdown("selectedGroupIdx", {
        label: strings.GroupPropertyFieldLabel,
        options: Utils.convertGrpToOptions(this.groups),
        selectedKey: this.properties.selectedGroupIdx,
      }),
      PropertyPaneTextField("teamsAppId", {
        label: strings.TeamsAppIdPropertyFieldLabel,
      }),
      PropertyFieldSitePicker("selectedSite", {
        label: "Select sites",
        initialSites: this.properties.selectedSite,
        context: this.context,
        deferredValidationTime: 500,
        multiSelect: false,
        onPropertyChange: this.onPropertyPaneFieldChanged,
        properties: this.properties,
        key: "sitesFieldId",
      }),
    ];

    if (this.properties?.selectedSite?.length) {
      this._groupsfields.push(
        PropertyFieldListPicker("selectedList", {
          label: "Select a list",
          selectedList: this.properties.selectedList,
          includeHidden: false,
          webAbsoluteUrl: this.properties.selectedSite[0]?.url,
          orderBy: PropertyFieldListPickerOrderBy.Title,
          disabled: false,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          context: this.context,
          onGetErrorMessage: null,
          deferredValidationTime: 0,
          key: "listPickerFieldId",
          includeListTitleAndUrl: true,
        })
      );
    }
  }

  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this._getGroupFields();
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: this._groupsfields,
            },
          ],
        },
      ],
    };
  }
}
