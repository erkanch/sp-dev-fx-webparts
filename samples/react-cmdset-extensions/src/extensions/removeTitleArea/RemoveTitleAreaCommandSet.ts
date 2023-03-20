/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/explicit-function-return-type */

import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import CustomPanel, { ICustomPanelProps } from './CustomPanel';
import { assign } from '@uifabric/utilities';
import * as React from 'react';
import * as ReactDom from 'react-dom';

const LOG_SOURCE: string = 'RemoveTitleAreaCommandSet';

export default class RemoveTitleAreaCommandSet extends BaseListViewCommandSet<{}> {

  private selectedFileName: string;
  private selectedFilePath: string;
  private panelPlaceHolder: HTMLDivElement = null;

  public onInit(): Promise<void> {
    this.context.listView.listViewStateChangedEvent.add(this, this._listViewStateChanged);

    let cmd: Command = this.tryGetCommand('COMMAND_ADD');
    cmd.visible = false;
    cmd = this.tryGetCommand('COMMAND_REMOVE');
    cmd.visible = false;
    cmd = this.tryGetCommand('COMMAND_MOVE');
    cmd.visible = false;
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
    this.dismissPanel = this.dismissPanel.bind(this);
    Log.info(LOG_SOURCE, 'Initialized RemoveTitleAreaCommandSet');
    return Promise.resolve();
  }

  private _listViewStateChanged(args: ListViewStateChangedEventArgs): void {

    const showCommand = (): boolean => {
      return this.context.listView.selectedRows?.length === 1 && this.context.listView.selectedRows[0].getValueByName("ContentType") !== "Folder";
    }

    if (this.context.pageContext.list?.title === "Site Pages") {
      const compareRemoveCommand: Command = this.tryGetCommand('COMMAND_REMOVE');
      if (compareRemoveCommand) {
        // This command should be hidden unless exactly one row is selected.
        compareRemoveCommand.visible = showCommand();
      }
      const compareAddCommand: Command = this.tryGetCommand('COMMAND_ADD');
      if (compareAddCommand) {
        // This command should be hidden unless exactly one row is selected.
        compareAddCommand.visible = showCommand();
      }
      const compareMoveCommand: Command = this.tryGetCommand('COMMAND_MOVE');
      if (compareMoveCommand) {
        // This command should be hidden unless exactly one row is selected.
        compareMoveCommand.visible = showCommand() && !this.context.listView.folderInfo.folderPath.endsWith("/Templates");
      }
    }
    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters) {

    const fileName = event.selectedRows[0].getValueByName("FileLeafRef");
    this.selectedFileName = fileName;
    this.selectedFilePath = event.selectedRows[0].getValueByName("FileRef");
    switch (event.itemId) {
      case 'COMMAND_REMOVE':
        this.updateLayout(event.selectedRows[0].getValueByName("ID"), "Home")
        .then(async u => {
          await Dialog.alert(`Removed title area from ${fileName}`);
            console.log("Updated", u);
            open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${fileName}`);
          })
          .catch(e => console.log("Failed", e));
        break;
      case 'COMMAND_ADD':
        this.updateLayout(event.selectedRows[0].getValueByName("ID"), "Article")
          .then(async u => {
            await Dialog.alert(`Added title area to ${fileName}`);
            console.log("Updated", u);
            open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${fileName}`);
          })
          .catch(e => console.log("Failed", e));
        break;
      case 'COMMAND_MOVE':
        await this.getSitePageFolderDetails(fileName);
        //this.moveItem();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private showPanel(currentTitle: string, folders) {
    this.renderPanelComponent({
      isOpen: true,
      currentTitle,
      onSave: this.onSave.bind(this),
      onClose: this.dismissPanel,
      onFolderClick: this.getChildFolders.bind(this),
      items: folders
    });
  }

  private async getChildFolders(folderName: string) {
    await this.getSitePageFolderDetails(this.selectedFileName, folderName);
  }

  private onSave(folderRelativeUrl: string) {
    if (!folderRelativeUrl) { // Move to root if empty
      folderRelativeUrl = `${this.context.pageContext.web.absoluteUrl.replace(location.origin, "")}/SitePages`;
    }
    this.moveItem(folderRelativeUrl)
      // eslint-disable-next-line @microsoft/spfx/no-async-await
      .then(async u => {
        await Dialog.alert(`Page moved successfully to ${folderRelativeUrl}`);
        location.reload();
        console.log("Updated", u);
      })
      .catch(e => console.log("Failed", e));
  }

  private dismissPanel() {
    this.renderPanelComponent({ isOpen: false });
  }

  private moveItem(folderRelativeUrl: string) {
    // return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/SP.MoveCopyUtil.MoveFileByPath()`,
    return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/getFileByServerRelativeUrl('${this.selectedFilePath}')/moveTo(newurl='${folderRelativeUrl}/${this.selectedFileName}',flags=1)`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
          'odata-version': '3.0'
        },
      });
  }

  private renderPanelComponent(props: unknown) {
    const element: React.ReactElement<ICustomPanelProps> = React.createElement(CustomPanel, assign({
      onClose: null,
      onSave: null,
      // onFolderClick: null,
      currentTitle: null,
      isOpen: false,
      items: null
    }, props));
    ReactDom.render(element, this.panelPlaceHolder);
  }

  private async getSitePageFolderDetails(fileName: string, folderName?: string) {
    let queryString = ""
    if (folderName)
      queryString = `/_api/Web/GetFolderByServerRelativeUrl('${this.context.pageContext.site.serverRelativeUrl}/SitePages/${folderName}')/Folders?$Select=Name,ServerRelativeUrl`;
    else
      queryString = `/_api/Web/GetFolderByServerRelativeUrl('${this.context.pageContext.site.serverRelativeUrl}/SitePages')/Folders?$Select=Name,ServerRelativeUrl`

    const folderQuery = await this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}${queryString}`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json'
        },
      });
    const folderItems = await folderQuery.json()
    const folders = folderItems?.value.filter((item) => item.Name !== "Templates" && item.Name !== "Forms");
    const count = (this.selectedFilePath.match(/\//g) || []).length;
    console.log(count);
    if (!folderName && count > 4) {
      folders.unshift({
        Name: "Root Folder",
        ServerRelativeUrl: `${this.context.pageContext.site.serverRelativeUrl}/SitePages`
      });
    }
    this.showPanel(fileName, folders);
  }

  private updateLayout(id: unknown, layout: string): Promise<SPHttpClientResponse> {

    const body: string = JSON.stringify({
      "__metadata":
      {
        "type": "SP.Data.SitePagesItem"
      },
      "PageLayoutType": layout
    });

    return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site Pages')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
          'odata-version': '3.0'
        },
        body
      });
  }

}
