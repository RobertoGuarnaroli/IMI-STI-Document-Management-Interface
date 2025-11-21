import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import * as React from "react";
import * as ReactDOM from "react-dom";
import DocumentManagementInterface from "./components/DocumentManagementInterface";
import { IDocumentManagementInterfaceProps } from "./components/IDocumentManagementInterfaceProps";

export interface IDocumentManagementInterfaceWebPartProps {
  description: string;
}

export default class DocumentManagementInterfaceWebPart extends BaseClientSideWebPart<IDocumentManagementInterfaceWebPartProps> {
  protected render(): void {
    const element: React.ReactElement<IDocumentManagementInterfaceProps> = React.createElement(
      DocumentManagementInterface, 
      { 
        context: this.context
      }
    );
    
    ReactDOM.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    // Inizializza PnPjs globalmente (opzionale)
    const sp = spfi().using(SPFx(this.context));
    (window as any).sp = sp;
  }
}