import * as React from "react";
import styles from "./Providerpart.module.scss";
import { IProviderpartProps } from "./IProviderpartProps";
import { IDepartment } from "./IDepartment";
import {
  CheckboxVisibility,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
} from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

let departmentListColumns = [
  {
    key: "ID",
    name: "ID",
    fieldName: "ID",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "Title",
    name: "Title",
    fieldName: "Title",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
];

const Providerpart: React.FC<IProviderpartProps> = (props) => {
  let _selection: Selection;

  //const [status, setStatus] = React.useState<string>("Ready");
  const [departmentListItems, setDepartmentListItems] = React.useState<
    IDepartment[]
  >([]);

  const [departmentItem, setDepartmentItem] = React.useState<IDepartment>({
    Id: 0,
    Title: "",
  });

  //const [mySelection, setMySelection] = React.useState();

  const onItemsSelectedChanged = () => {
    // props.onDepartmentSelected(_selection.getSelection()[0] as IDepartment);
    // setDepartmentItem(_selection.getSelection()[0] as IDepartment);
    console.log("DEPT1 " + JSON.stringify(departmentItem));
    setDepartmentItem(_selection.getSelection()[0] as IDepartment);
    props.onDepartmentSelected(departmentItem);
    console.log("DEPT2 " + JSON.stringify(departmentItem));
  };

  React.useEffect(() => {
    getListItems();
  }, []);

  React.useEffect(() => {
    _selection = new Selection({
      onSelectionChanged: onItemsSelectedChanged,
    });
  }, [_selection]);

  const getListItems = () => {
    props.context.spHttpClient
      .get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Departments')/items`,
        SPHttpClient.configurations.v1
      )
      .then((res: SPHttpClientResponse): Promise<{ value: IDepartment[] }> => {
        return res.json();
      })
      .then((resp) => setDepartmentListItems(resp.value));
  };

  return (
    <section className={`${styles.providerpart}`}>
      <DetailsList
        items={departmentListItems}
        columns={departmentListColumns}
        setKey="Id"
        checkboxVisibility={CheckboxVisibility.always}
        selectionMode={SelectionMode.single}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        compact={true}
        selection={_selection}
      />
    </section>
  );
};

export default Providerpart;
