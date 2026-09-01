export interface IPeopleSearchWebPartProps {
  placeholderText: string;
  height: number;
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
