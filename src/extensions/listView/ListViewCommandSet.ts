import { override } from "@microsoft/decorators";
import { Log, DisplayMode } from "@microsoft/sp-core-library";
import * as SP from "./SPOHelper";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  BaseFieldCustomizer,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "ListViewCommandSetStrings";
import DialogManager from "@microsoft/sp-dialog/lib/DialogManager";
import { DialogState } from "@microsoft/sp-dialog/lib/BaseDialog";
import ColorPickerDialog from "./CustomDialogBox";
import { IColor, IDropdown } from "office-ui-fabric-react";
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListViewCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "ListViewCommandSet";

export default class ListViewCommandSet extends BaseListViewCommandSet<
  IListViewCommandSetProperties
> {
  private _colorCode: IColor;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized ListViewCommandSet");
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length > 0;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_1":
        let csvcontent = "data:text/csv;charset=utf-8,";
        event.selectedRows.map((item) => {
          const cases = item.getValueByName("ConfirmedCases");
          const val =
            item.getValueByName("Title") +
            "," +
            cases.replace(",", "") +
            "," +
            item.getValueByName("ID") +
            "\n";
          csvcontent += val;
        });
        var encodedUri = encodeURI(csvcontent);
        var link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "selectedRecords.csv");
        document.body.appendChild(link);
        link.click();

        Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      case "COMMAND_2":
        SP.SPGet(
          `https://gayasagar.sharepoint.com/sites/MySite/_api/web/lists/getbytitle('Cases')/items`
        ).then((r) => {
          let csvcontent = "data:text/csv;charset=utf-8,";
          r.value.map((item) => {
            const val =
              item.Title + "," + item.ConfirmedCases + "," + item.ID + "\n";
            csvcontent += val;
          });
          var encodedUri = encodeURI(csvcontent);
          var link = document.createElement("a");
          link.setAttribute("href", encodedUri);
          link.setAttribute("download", "my_data.csv");
          document.body.appendChild(link);
          link.click();
        });
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      case "COMMAND_3":
        const dialog: ColorPickerDialog = new ColorPickerDialog();
        dialog.message = "Select Fields to export:";
        // Use 'FFFFFF' as the default color for first usage
        dialog.show().then(() => {
          let values = "";
          dialog.seletedFields.map((item) => (values += item.key + ","));
          const fields = values.substring(0, values.length - 1);
          SP.SPGet(
            `https://gayasagar.sharepoint.com/sites/MySite/_api/web/lists/getbytitle('Cases')/items?$select=${fields}`
          ).then((r) => {
            let csvcontent = "data:text/csv;charset=utf-8,";
            r.value.map((item) => {
              let val = "";
              Object.keys(item).forEach((key) => {
                val += item[key] + ",";
              });
              csvcontent += val.substring(0, val.length - 1) + "\n";
            });
            console.log(csvcontent);
            var encodedUri = encodeURI(csvcontent);
            var link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", "selectedFields.csv");
            document.body.appendChild(link);
            link.click();
          });
          Dialog.alert(`Your File has been dowloading..`);
        });
        break;
      default:
        throw new Error("Unknown command");
    }
  }
}
