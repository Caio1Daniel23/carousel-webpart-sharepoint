export interface IPeopleSearchWebPartProps {
  placeholderText: string;
  height: number;
  showHeader: boolean;
  headerText: string;
  searchLabelText: string;
}

export interface IPersonResult {
  id: string;
  displayName: string;
  mail?: string;
  jobTitle?: string;
  businessPhones?: string[];
  officeLocation?: string;
  department?: string;
}
