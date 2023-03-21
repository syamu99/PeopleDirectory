import { IPerson } from "../PeopleDirectory";

export interface IPeopleListState {
  showCallOut: boolean;
  calloutElement: number;
  person: IPerson;
  paginatedItems:IPerson[];
  allItems:IPerson[];
}