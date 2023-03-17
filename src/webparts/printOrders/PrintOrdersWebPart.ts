import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "PrintOrdersWebPartStrings";
import PrintOrders from "./components/PrintOrders";
import { IPrintOrdersProps } from "./components/IPrintOrdersProps";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

export interface IPrintOrdersWebPartProps {
  description: string;
  DatosAI: any;
}

export default class PrintOrdersWebPart extends BaseClientSideWebPart<IPrintOrdersWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IPrintOrdersProps> = React.createElement(
      PrintOrders,
      {
        description: this.properties.description,
        context: this.context,
        DatosAI: this.properties.DatosAI,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      spfi().using(SPFx(this.context));
    });
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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker("DatosAI", {
                  label: "Selecciona la lista de DatosAI",
                  selectedList: this.properties.DatosAI,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
