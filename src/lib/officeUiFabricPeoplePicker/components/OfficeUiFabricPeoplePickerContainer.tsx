import * as React from "react";
import { IOfficeUiFabricPeoplePickerContainerProps } from "./IOfficeUiFabricPeoplePickerContainerProps";
import {
  CompactPeoplePicker,
  NormalPeoplePicker
} from "office-ui-fabric-react/lib/Pickers";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
import { people } from "./PeoplePickerExampleData";
import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { SharePointUserPersona } from "../models/OfficeUiFabricPeoplePicker";
import {
  sp,
  PeoplePickerEntity,
  ClientPeoplePickerQueryParameters,
  WebEnsureUserResult,
  stringIsNullOrEmpty
} from "@pnp/pnpjs";
import { IOfficeUiFabricPeoplePickerContainerState } from "./IOfficeUiFabricPeoplePickerContainerState";
import { OfficeUiFabricPeoplePicker } from "./OfficeUiFabricPeoplePicker";

export class OfficeUiFabricPeoplePickerContainer extends React.Component<
  IOfficeUiFabricPeoplePickerContainerProps,
  IOfficeUiFabricPeoplePickerContainerState
> {
  constructor(props) {
    super(props);
    this.state = {
      selectedItems: undefined
    };
  }

  public render(): React.ReactElement<
    IOfficeUiFabricPeoplePickerContainerProps
  > {
    return (
      <OfficeUiFabricPeoplePicker
        typePicker={this.props.typePicker}
        onResolveSuggestions={this._onFilterChanged}
        selectedItems={this.state.selectedItems}
        onChange={this._onSelectionChange}
      />
    );
  }

  public componentDidMount() {
    if (Environment.type === EnvironmentType.Local) {
    } else {
      this._fetchDefaultSelection(this.props.defaultSelectionEmails);
    }
  }

  private _getSharePointUserPersonas(
    entities: PeoplePickerEntity[]
  ): Promise<SharePointUserPersona[]> {
    var batch = sp.web.createBatch();

    let personas = [];

    entities.map((entity: PeoplePickerEntity) => {
      batch.addResolveBatchDependency(
        sp.web
          .inBatch(batch)
          .ensureUser(entity.EntityData.Email)
          .then((result: WebEnsureUserResult) => {
            personas.push(new SharePointUserPersona(entity, result));
          })
          .catch((error: any) => {
            console.log(error);
          })
      );
    });

    return batch.execute().then(_ => {
      return personas;
    });
  }

  private _fetchDefaultSelection = (emails: string[]) => {
    var batch = sp.web.createBatch();

    let entities: PeoplePickerEntity[] = [];

    emails.map((email: string) => {
      batch.addResolveBatchDependency(
        sp.profiles
          .clientPeoplePickerSearchUser(this._getQueryParams(email))
          .then((result: PeoplePickerEntity[]) => {
            if (result.length === 1) {
              entities.push(result[0]);
            } else {
              console.log("multiple entities fetched");
            }
          })
      );
    });

    batch.execute().then(_ => {
      this._getSharePointUserPersonas(entities).then(
        (personas: SharePointUserPersona[]) => {
          this.setState({ selectedItems: personas });
        }
      );
    });
  }

  private _onSelectionChange = (selections: SharePointUserPersona[]) => {
    this.setState(
      {
        selectedItems: selections
      },
      () => {
        let selectionIds: number[] = [];
        selections.map((selection: SharePointUserPersona) => {
          selectionIds.push(selection.user.Id);
        });
        this.props.onChange(selectionIds);
      }
    );
  }

  private _onFilterChanged = (filterText: string) => {
    if (stringIsNullOrEmpty(filterText)) {
      return [];
    } else {
      if (filterText.length > 2) {
        return this._searchPeople(filterText);
      }
    }
  }

  private _getQueryParams(terms: string): ClientPeoplePickerQueryParameters {
    let principalType: number = 0;
    if (this.props.principalTypeUser === true) {
      principalType += 1;
    }
    if (this.props.principalTypeSharePointGroup === true) {
      principalType += 8;
    }
    if (this.props.principalTypeSecurityGroup === true) {
      principalType += 4;
    }
    if (this.props.principalTypeDistributionList === true) {
      principalType += 2;
    }
    return {
      AllowEmailAddresses: true,
      AllowMultipleEntities: false,
      AllUrlZones: false,
      MaximumEntitySuggestions: this.props.numberOfItems,
      PrincipalSource: 15,
      // PrincipalType controls the type of entities that are returned in the results.
      // Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
      // These values can be combined (example: 13 is security + SP groups + users)
      PrincipalType: principalType,
      QueryString: terms
    };
  }

  /**
   * @function
   * Returns fake people results for the Mock mode
   */
  private searchPeopleFromMock(terms: string): SharePointUserPersona[] {
    return people.filter((value: SharePointUserPersona) => {
      if (value.text.toLowerCase().indexOf(terms.toLowerCase()) !== -1) {
        return value;
      }
    });
  }

  /**
   * @function
   * Returns people results after a REST API call
   */
  private _searchPeople(
    terms: string
  ): SharePointUserPersona[] | Promise<SharePointUserPersona[]> {
    if (Environment.type === EnvironmentType.Local) {
      // If the running environment is local, load the data from the mock
      return this.searchPeopleFromMock(terms);
    } else {
      return sp.profiles
        .clientPeoplePickerSearchUser(this._getQueryParams(terms))
        .then((entities: PeoplePickerEntity[]) => {
          return this._getSharePointUserPersonas(entities);
        })
        .catch((error: any) => {
          console.log(error);
          return [];
        });
    }
  }
}
