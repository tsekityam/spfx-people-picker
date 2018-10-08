export interface IOfficeUiFabricPeoplePickerContainerProps {
  typePicker: string;
  principalTypeUser: boolean;
  principalTypeSharePointGroup: boolean;
  principalTypeSecurityGroup: boolean;
  principalTypeDistributionList: boolean;
  numberOfItems: number;
  onChange?: (selectionIds: number[]) => void;
  defaultSelectionEmails?: string[];
}
