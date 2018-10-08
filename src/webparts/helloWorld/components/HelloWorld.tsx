import * as React from "react";
import { IHelloWorldProps } from "./IHelloWorldProps";
import {
  OfficeUiFabricPeoplePicker,
  SharePointUserPersona
} from "../../../lib/officeUiFabricPeoplePicker";
import { IHelloWorldState } from "./IHelloWorldState";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { sp } from "@pnp/sp";

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorldState
> {
  constructor(props) {
    super(props);
    this.state = {
      defaultSelectedItems: []
    };
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <OfficeUiFabricPeoplePicker
        typePicker={"Normal"}
        principalTypeUser={true}
        principalTypeSharePointGroup={false}
        principalTypeSecurityGroup={false}
        principalTypeDistributionList={false}
        numberOfItems={5}
        onChange={(selectionIds: number[]) => {
          console.log(selectionIds);
        }}
        defaultSelectionEmails={["tseki@esquel.com", "qiuzx@esquel.com"]}
      />
    );
  }

  public componentDidMount() {
    if (Environment.type === EnvironmentType.Local) {
    } else {
      // sp.web.lists
      //   .getByTitle("Promotions")
      //   .items.select("Editor/EMail")
      //   .expand("Editor")
      //   .getAll()
      //   .then((result: any[]) => {
      //     console.log(result);
      //   });
    }
  }
}
