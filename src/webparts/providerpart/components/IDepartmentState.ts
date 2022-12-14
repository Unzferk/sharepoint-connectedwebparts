import { IDepartment } from "./IDepartment";

export interface IDepartmentState {
  status: string;
  DepartmentListItems: IDepartment[];
  DepartmentListItem: IDepartment;
}
