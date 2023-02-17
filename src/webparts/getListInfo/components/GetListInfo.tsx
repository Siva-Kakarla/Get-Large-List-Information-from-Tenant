import * as React from 'react';
import styles from './GetListInfo.module.scss';
import { IGetListInfoProps } from './IGetListInfoProps';
import ListinfoComp from './ListInfoComp/ListInfoComp';
// import { escape } from '@microsoft/sp-lodash-subset';

// import interfaces
//import { IFile } from "./interfaces";//,IResponseItem

// import { Caching } from "@pnp/queryable";
// import { getSP } from "../pnpjsConfig";
// import { SPFI, spfi } from "@pnp/sp";
// import { Logger, LogLevel } from "@pnp/logging";
// import { IItemUpdateResult } from "@pnp/sp/items";

// import { graph } from "@pnp/graph";

import {MSGraphClientV3} from '@microsoft/sp-http';


import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
//import {DetailsList} from 'office-ui-fabric-react';

// import { IODataUser, IODataWeb } from '@microsoft/sp-odata-types';



import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
//import { ISPHttpClientBatchOptions, SPHttpClientBatch, ISPHttpClientBatchCreationOptions } from '@microsoft/sp-http';

//import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

//import { getSP } from '../pnpjsConfig';
//import { Web } from '@pnp/sp/webs';
//import { result } from 'lodash';


import {Modal} from 'react-bootstrap';

export interface ISpSite {
  displayName: string;
  webUrl: string;
  id: string;
}

export interface IListInfo {
  name: string;
  displayName: string;
  webUrl: string;
  siteUrl: string;
  id: string;
  hidden: boolean;
  contentTypesEnabled: boolean;
}

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

export interface ISpSiteState {
  ListOf_SpSites:ISpSite[];
  ListOf_Lists:IListInfo[];
  ListOf_Final_ListInfo: IFinal_ListInfo[];
  ShowPopUp: boolean;
}


export default class GetListInfo extends React.Component<IGetListInfoProps, ISpSiteState> {

  // private _sp: SPFI;

  public all_SpSites: ISpSite[] = [];
  public all_ListInfo: IListInfo[] = [];
  public all_Final_ListInfo: IFinal_ListInfo[] = [];
  //private _sp: SPFI;

  constructor(props: IGetListInfoProps) {
    super(props);
    this.state = {
      ListOf_SpSites: [],
      ListOf_Lists: [],
      ListOf_Final_ListInfo: [],
      ShowPopUp: false
    };
    //this._sp = getSP();
  }

  public componentDidMount(): void {
    // read all file sizes from Documents library
    //this._Get_Site_Collection();//_readAllFilesSize();
    
  }

  public render() : React.ReactElement<IGetListInfoProps> {

    const _ListInfoCmp : JSX.Element = ( (this.state.ListOf_Final_ListInfo.length > 0) ?  <ListinfoComp items={this.state.ListOf_Final_ListInfo}/> : null);
    
    const _modelPopUp : JSX.Element = ( 
    <div className="modal show" style={{ display: 'block', position: 'initial' }}>
        <Modal
            show={this.state.ShowPopUp}
            backdrop="static"
            keyboard={false}
            size="lg"
            aria-labelledby="contained-modal-title-vcenter"
            centered>

        <Modal.Header>
            <Modal.Title>Please Wait</Modal.Title>
        </Modal.Header>

        <Modal.Body>
            <p>Please Wait we are finding the lists..</p>
            <img src="https://i.stack.imgur.com/hzk6C.gif" alt="Girl in a jacket" width="auto" height="auto"></img>
        </Modal.Body>

        <Modal.Footer>
            
        </Modal.Footer>
        </Modal>
    </div>)
    return (
      <div className={styles.getListInfo}>
          
          <h3>Find Large Lists In Your Tenant</h3>
          <PrimaryButton text="Begin Search.." onClick={this._Get_Site_Collection}/>
          
          {/* <PrimaryButton text="Search Lists" onClick={this._Get_List_Details}>Update Item Titles</PrimaryButton> */}
          {/* <DetailsList items={this.state.ListOf_SpSites}/> */}

          {_ListInfoCmp}
          {_modelPopUp}

        </div>
    );
  }

  public handleClose = () => {
    this.setState({ShowPopUp:false});
  }

  public _resetVariables = () =>{
    this.all_SpSites = [];
    this.all_ListInfo = [];
    this.all_Final_ListInfo = [];
  }

  /*public _PnPJsCheck = async (List_obj:any) => {
    // const item: any = await this._sp.web.lists.getByTitle("My List").items.getById(1)();
    // console.log(item);

    console.log(this._sp);

    const currentWebUrl: string = List_obj.siteUrl;

    const spWebA = spfi().using(SPFx(this.props.context));
    const spWebE = Web([spWebA.web, currentWebUrl]);
    debugger;
    let ItemCount = 0;
    ItemCount = await spWebE.lists.getByTitle(List_obj.name).items.length;
    // await spWebE.lists.getByTitle(List_obj.name).get().then((result)=>{
    //   ItemCount = result.ItemCount
    // });//"'"+  +"'"
    console.log(ItemCount);
  }*/

  //https://graph.microsoft.com/v1.0/sites/
  
  private _Get_Site_Collection = async () => {
    
    this._resetVariables();
    this.setState({ ShowPopUp:true });
    
    await this.props.context.msGraphClientFactory.getClient('3').then(async (msGraphClient: MSGraphClientV3) =>{
     await msGraphClient.api("sites").version("v1.0").select("displayName,name,webUrl,id").query("search=*").get(async (err: any,res:any)=>{
        if(err){
            console.log("Error Occured", err);
        }
        
        res.value.map((results: any)=>{
          this.all_SpSites.push({
            displayName: results.displayName,
            webUrl: results.webUrl,
            id: results.id
          })
        });

        //Check for Subsites in the Site_Collections//

        for(let i=0;i<=this.all_SpSites.length-1; i++){
            await this._Get_Sub_Sites(this.all_SpSites[i].id);
        }
        this.setState({ListOf_SpSites: this.all_SpSites});

        //Check for List in all the Site_Coll & Sub_Sites//

        for(let i=0;i<=this.state.ListOf_SpSites.length-1; i++){
            await this._Get_Lists(this.state.ListOf_SpSites[i]);
        }
        this.setState({ListOf_Lists: this.all_ListInfo});

        for(let i=0;i<=this.state.ListOf_Lists.length-1; i++){
            await this._Get_List_Item_Count(this.state.ListOf_Lists[i]);
            //await this._PnPJsCheck(this.state.ListOf_Lists[i]);
        }
        this.setState({ListOf_Final_ListInfo: this.all_Final_ListInfo});

        this.handleClose();
      })
    });
  };

  //https://graph.microsoft.com/v1.0/sites/{site-id}/sites
  private _Get_Sub_Sites = async (SiteCollectionID : string) => {//, MaxCount: number) => {
    await this.props.context.msGraphClientFactory.getClient('3').then(async (msGraphClient: MSGraphClientV3) =>{
      await msGraphClient.api("/sites/"+ SiteCollectionID +"/sites").version("v1.0").select("displayName,name,webUrl,id").get((err: any,res:any)=>{
        if(err){
            console.log("Error Occured", err);
        }
        res.value.map((results: any)=>{
          this.all_SpSites.push({
            displayName: results.displayName,
            webUrl: results.webUrl,
            id: results.id
          })
        });
      })
    });
  };

  //https://graph.microsoft.com/v1.0/sites/{site-id}/lists
  private _Get_Lists = async (SiteCollectionID : any) => {
    await this.props.context.msGraphClientFactory.getClient('3').then(async (msGraphClient: MSGraphClientV3) =>{
      await msGraphClient.api("/sites/"+ SiteCollectionID.id+"/lists").version("v1.0").select("displayName,name,webUrl,id,list").get((err: any,res:any)=>{
        if(err){
            console.log("Error Occured", err);
        }
        debugger;
        res.value.map((results: any)=>{
          this.all_ListInfo.push({
            name: results.name,
            displayName: results.displayName,
            webUrl: results.webUrl,
            siteUrl: SiteCollectionID.webUrl,
            id: results.id,
            hidden: results.list.hidden,
            contentTypesEnabled: results.list.contentTypesEnabled
          })
        });
      })
    });
  };

  private _Get_List_Item_Count = async (List_obj : any) => {

    debugger;
    const spHttpClient: SPHttpClient = this.props.context.spHttpClient;
    const currentWebUrl: string = List_obj.siteUrl;


    // Need to change the URL of the Site -> Where the List exists.
    //const currentWebUrl: string = this.props.context.pageContext.web.absoluteUrl;

    //GET 

    try{
      await spHttpClient.get(`${currentWebUrl}` + "/_api/web/lists/getbytitle('"+ List_obj.name +"')/ItemCount", SPHttpClient.configurations.v1).then(async (response: SPHttpClientResponse) => {
        await response.json().then((rsp: any) => {
                console.log(rsp);
                if(rsp.error)
                {
                    console.log("Error Occured", rsp.error);
                }
                else
                {
                  if(rsp.value >0)
                  {
                    this.all_Final_ListInfo.push({
                      Name: List_obj.name,
                      DisplayName: List_obj.displayName,
                      Count: rsp.value,
                      Hidden: List_obj.hidden,
                      ListUrl: List_obj.webUrl,
                      SiteUrl: List_obj.siteUrl,
                      Id: List_obj.id,
                      ContentTypesEnabled: List_obj.contentTypesEnabled
                    })
                  }
                }
        });
      });
    }
    catch(ex)
    {
      debugger;
      console.log(ex);
    }
  }

}
