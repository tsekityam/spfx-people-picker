import * as React from "react";
import { IOfficeUiFabricPeoplePickerProps } from "./IOfficeUiFabricPeoplePickerProps";
import {
  CompactPeoplePicker,
  NormalPeoplePicker
} from "office-ui-fabric-react/lib/Pickers";

export class OfficeUiFabricPeoplePicker extends React.Component<
  IOfficeUiFabricPeoplePickerProps,
  {}
> {
  public render(): React.ReactElement<IOfficeUiFabricPeoplePickerProps> {
    if (this.props.typePicker == "Normal") {
      return (
        <NormalPeoplePicker
          onChange={this.props.onChange}
          onResolveSuggestions={this.props.onResolveSuggestions}
          selectedItems={this.props.selectedItems}
        />
      );
    } else {
      return (
        <CompactPeoplePicker
          onChange={this.props.onChange}
          onResolveSuggestions={this.props.onResolveSuggestions}
          selectedItems={this.props.selectedItems}
        />
      );
    }
  }
}
