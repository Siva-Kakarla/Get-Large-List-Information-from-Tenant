import * as React from 'react';
import styles from './ListInfoComp.module.scss';

import {Table} from 'react-bootstrap';
import 'bootstrap/dist/css/bootstrap.min.css';

import { CommandBarButton } from 'office-ui-fabric-react/lib/Button';
import { CSVLink } from "react-csv";

export interface IFinal_ListInfo {
    Name: string;
    DisplayName: string;  
    Count: number;
    Hidden: boolean;
    ListUrl: string;
    SiteUrl: string;
    Id: string;
    ContentTypesEnabled: boolean;
  }

export interface IListInfoCompProps{
    items: IFinal_ListInfo[]
}

export default class ListinfoComp extends React.Component<IListInfoCompProps, {}> {

  constructor(props: IListInfoCompProps)
  {
    super(props);
  }

  public render() : React.ReactElement<IListInfoCompProps> {

    let _table_Body : JSX.Element = ((this.props.items.length > 0) ? <tbody>
    {this.props.items.map( element => 
      <tr>
        <td>{element.Name}</td>
        <td>{element.DisplayName}</td>
        <td>{element.Count}</td>
        <td>{element.Hidden.toString()}</td>
        <td>{element.ListUrl}</td>
        <td>{element.SiteUrl}</td>
        <td>{element.Id}</td>
        <td>{element.ContentTypesEnabled.toString()}</td>
      </tr>
    )}      
  </tbody> : null);

    return (
      <div className={styles.ListInfoComp}>
          
          <h3>Large List Info: </h3>

          <div>
            <CSVLink data={this.props.items} filename={'LargeListInfo.csv'}>
                <CommandBarButton  iconProps={{ iconName: 'ExcelLogoInverse' }} text='Export to Excel' />
            </CSVLink>
          </div>

          <Table striped bordered hover size="sm" responsive>
            <thead>
              <tr>
                <th>Name</th>
                <th>Display Name</th>
                <th>Count</th>                
                <th>Is Hidden List</th>
                <th>List Url</th>
                <th>Site Url</th>
                <th>List Id</th>
                <th>ContentTypesEnabled</th>
              </tr>
            </thead>

            {_table_Body}
          </Table>
          
        </div>
    );
  }

}
