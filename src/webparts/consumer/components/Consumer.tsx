import * as React from "react";
import styles from "./Consumer.module.scss";
import { ICostumerProps } from "./IConsumerProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
} from "office-ui-fabric-react";
import { IEmployee } from "./IEmployee";

let _employeeColumns = [
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
  {
    key: "DeptTitle",
    name: "DeptTitle",
    fieldName: "DeptTitleId",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
];

const Costumer: React.FC<ICostumerProps> = (props) => {
  //const [status, setStatus] = React.useState();
  const [deptTitleId, setDeptTitleId] = React.useState("");
  const [employeesListItems, setEmployeesListItems] = React.useState([]);
  /* const [employeeItem, setEmployeeItem] = React.useState({
    Id: 0,
    Title: "",
    Dpt: "",
  });*/

  // React.useEffect(() => {
  //   getListItems();
  // }, [props.DeptTitleId]);

  const getListItems = () => {
    props.context.spHttpClient
      .get(
        `${
          props.context.pageContext.web.absoluteUrl
          //}/_api/web/lists/getbytitle('Employees')/items?$filter=DeptTitleId eq ${props.DeptTitleId.tryGetValue()}`,
        }/_api/web/lists/getbytitle('Employees')/items?&$filter=DeptTitleId eq ${props.DeptTitleId.tryGetValue()}`,
        SPHttpClient.configurations.v1
      )
      .then((res: SPHttpClientResponse): Promise<{ value: IEmployee[] }> => {
        return res.json();
      })
      .then((resp) => {
        console.log("RESP: " + JSON.stringify(resp));

        setEmployeesListItems(resp.value);
        setDeptTitleId(props.DeptTitleId.tryGetValue().toString());
        console.log(deptTitleId);
      });
  };

  return (
    <section className={`${styles.consumer}`}>
      <h1>Selected Department is: {props.DeptTitleId.tryGetValue()}</h1>
      <DetailsList
        items={employeesListItems}
        columns={_employeeColumns}
        setKey="Id"
        checkboxVisibility={CheckboxVisibility.always}
        selectionMode={SelectionMode.single}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        compact={true}
      />
      <button onClick={() => getListItems()}>fetch</button>
    </section>
  );
};

export default Costumer;
