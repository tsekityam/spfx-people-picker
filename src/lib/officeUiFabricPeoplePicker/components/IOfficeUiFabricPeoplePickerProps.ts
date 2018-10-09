import { SharePointUserPersona } from "../models";

export interface IOfficeUiFabricPeoplePickerProps {
  typePicker: string;
  onResolveSuggestions: (
    filter: string,
    selectedItems?: SharePointUserPersona[]
  ) => SharePointUserPersona[] | PromiseLike<SharePointUserPersona[]>;
  selectedItems?: SharePointUserPersona[];
  onChange?: (selectedItems?: SharePointUserPersona[]) => void;
}
