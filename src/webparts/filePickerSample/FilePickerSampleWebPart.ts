import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "FilePickerSampleWebPartStrings";
import { ListRepository } from "../../repositories/List/ListRepository";
import {
  FilePickerSample,
  IFilePickerSampleProps,
} from "./components/FilePickerSample";

export interface IFilePickerSampleWebPartProps {
  imageStorageSharePointDocumentLibrary: string;
}

export default class FilePickerSampleWebPart extends BaseClientSideWebPart<IFilePickerSampleWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];

  public render(): void {
    const element: React.ReactElement<IFilePickerSampleProps> =
      React.createElement(FilePickerSample, {
        wpp: this.properties,
        context: this.context,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "lists"
    );

    const listRespository = new ListRepository(this.context);

    listRespository
      .getLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
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
                PropertyPaneDropdown("imageStorageSharePointDocumentLibrary", {
                  label:
                    "Select SharePoint document library for custom image upload",
                  options: this.lists,
                  selectedKey:
                    this.properties.imageStorageSharePointDocumentLibrary,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
