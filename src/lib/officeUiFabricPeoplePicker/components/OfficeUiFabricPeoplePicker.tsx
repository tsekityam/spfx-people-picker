import * as React from "react";
import { IOfficeUiFabricPeoplePickerProps } from "./IOfficeUiFabricPeoplePickerProps";
import {
  CompactPeoplePicker,
  IBasePickerSuggestionsProps,
  NormalPeoplePicker
} from "office-ui-fabric-react/lib/Pickers";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
import { people } from "./PeoplePickerExampleData";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { Promise } from "es6-promise";
import * as lodash from "lodash";
import {
  IOfficeUiFabricPeoplePickerState,
  SharePointUserPersona
} from "../models/OfficeUiFabricPeoplePicker";
import {
  sp,
  PeoplePickerEntity,
  ClientPeoplePickerQueryParameters,
  WebEnsureUserResult
} from "@pnp/pnpjs";

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggested People",
  noResultsFoundText: "No results found",
  loadingText: "Loading"
};

export class OfficeUiFabricPeoplePicker extends React.Component<
  IOfficeUiFabricPeoplePickerProps,
  IOfficeUiFabricPeoplePickerState
> {
  constructor() {
    super();
    this.state = {
      currentPicker: 1,
      delayResults: false,
      selectedItems: []
    };
  }

  public render(): React.ReactElement<IOfficeUiFabricPeoplePickerProps> {
    if (this.props.typePicker == "Normal") {
      return (
        <NormalPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
          pickerSuggestionsProps={suggestionProps}
          className={"ms-PeoplePicker"}
          key={"normal"}
        />
      );
    } else {
      return (
        <CompactPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
          pickerSuggestionsProps={suggestionProps}
          className={"ms-PeoplePicker"}
          key={"normal"}
        />
      );
    }
  }

  private _onChange(items: any[]) {
    this.setState({
      selectedItems: items
    });
    if (this.props.onChange) {
      this.props.onChange(items);
    }
  }

  private _onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[]
  ) => {
    if (filterText) {
      if (filterText.length > 2) {
        return this._searchPeople(filterText);
      }
    } else {
      return [];
    }
  }

  /**
   * @function
   * Returns fake people results for the Mock mode
   */
  private searchPeopleFromMock(terms: string): IPersonaProps[] {
    return people.filter((value: IPersonaProps) => {
      if (value.primaryText.toLowerCase().indexOf(terms.toLowerCase()) !== -1) {
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
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      // If the running environment is local, load the data from the mock
      return this.searchPeopleFromMock(terms);
    } else {
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
      const queryParams: ClientPeoplePickerQueryParameters = {
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

      return sp.profiles
        .clientPeoplePickerSearchUser(queryParams)
        .then((entities: PeoplePickerEntity[]) => {
          var batch = sp.web.createBatch();

          let personas = [];

          entities.map((entity: PeoplePickerEntity) => {
            sp.web
              .inBatch(batch)
              .ensureUser(entity.Key)
              .then((result: WebEnsureUserResult) => {
                personas.push(new SharePointUserPersona(entity, result));
              });
          });

          return batch.execute().then(_ => {
            console.log(personas);
            return personas;
          });
        });
    }
  }
}
