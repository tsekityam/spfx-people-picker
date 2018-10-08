import * as React from "react";
import styles from "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  OfficeUiFabricPeoplePicker,
  SharePointUserPersona
} from "../../../lib/officeUiFabricPeoplePicker";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <OfficeUiFabricPeoplePicker
        description={this.props.description}
        spHttpClient={this.props.spHttpClient}
        siteUrl={this.props.siteUrl}
        typePicker={"Normal"}
        principalTypeUser={true}
        principalTypeSharePointGroup={false}
        principalTypeSecurityGroup={false}
        principalTypeDistributionList={false}
        numberOfItems={5}
        onChange={(items: SharePointUserPersona[]) => {
          console.log(items);
        }}
      />
    );
  }
}
