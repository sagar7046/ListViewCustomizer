import * as React from "react";
import * as ReactDOM from "react-dom";
import * as SP from "./SPOHelper";
import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import {
  ColorPicker,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent,
  IColor,
  Dropdown,
  IDropdownOption,
} from "office-ui-fabric-react";

interface IColorPickerDialogContentProps {
  message: string;
  close: () => void;
  submit: (color: IDropdownOption[]) => void;
}

class CustomDialogBox extends React.Component<
  IColorPickerDialogContentProps,
  {}
> {
  private _pickedColor: IColor;

  constructor(props) {
    super(props);
    // Default Color
  }

  state = {
    options: [],
    selectedKeys: [],
  };
  componentDidMount() {
    SP.SPGet(
      `https://gayasagar.sharepoint.com/sites/MySite/_api/web/lists/GetByTitle('Cases')/fields?$filter=Hidden eq false`
    ).then((r) => {
      const option = [];
      r.value.map((item) => {
        option.push({
          key: item.EntityPropertyName,
          text: item.InternalName,
        });
      });
      this.setState({ options: option });
    });
  }

  _onChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    const fields = this.state.selectedKeys;
    if (item) {
      var result;
      if (item.selected == false) {
        result = fields.filter((i) => i.key !== item.key);
        this.setState({ selectedKeys: result }, () => {
          console.log(this.state.selectedKeys);
        });
      } else {
        fields.push(item);
        this.setState({ selectedKeys: fields }, () => {
          console.log(this.state.selectedKeys);
        });
      }
    }
  };

  public render(): JSX.Element {
    return (
      <DialogContent
        title="Color Picker"
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        <Dropdown
          options={this.state.options}
          multiSelect
          onChange={this._onChange}
          placeholder="Select options"
        ></Dropdown>
        <DialogFooter>
          <Button text="Cancel" title="Cancel" onClick={this.props.close} />
          <PrimaryButton
            text="OK"
            title="OK"
            onClick={() => {
              this.props.submit(this.state.selectedKeys);
            }}
          />
        </DialogFooter>
      </DialogContent>
    );
  }
}

export default class ColorPickerDialog extends BaseDialog {
  public message: string;
  public seletedFields = [];

  public render(): void {
    ReactDOM.render(
      <CustomDialogBox
        close={this.close}
        message={this.message}
        submit={this._submit}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false,
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();

    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  private _submit = (selectedFields) => {
    this.seletedFields = selectedFields.filter(
      (item) => item.selected === true
    );
    console.log(this.seletedFields);
    this.close();
  };
}
