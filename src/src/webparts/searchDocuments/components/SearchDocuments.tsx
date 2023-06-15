import * as React from 'react';
import styles from './SearchDocuments.module.scss';
import { ISearchDocumentsProps } from './ISearchDocumentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPInstance } from '../../../service';
import { getSP } from '../../../opspnp';
import { DetailsList, DetailsRow, IColumn, IDetailsListProps, IDetailsRowStyles, TextField } from 'office-ui-fabric-react';
import * as _ from 'lodash';

export default class SearchDocuments extends React.Component<ISearchDocumentsProps, any> {

  private ListServeInstance: SPInstance;
  private _columns: IColumn[];

  constructor(props: ISearchDocumentsProps) {
    super(props);
    this.state = {
      items: [],
      filteredItems: []
    }

    this._columns = [
      { key: 'title', name: 'Title', fieldName: 'LinkFilename', minWidth: 100, maxWidth: 200, isResizable: true,
      onRender: (item: any) => {
        var hrefValue =item.FileSystemObjectType == 0 ? item.FileRef + "?web=1" : item.FileRef;
        if(item.FileSystemObjectType == 1){
          return <><a data-intercenption="off" target="_blank" id="anchorStyleFolder" href={hrefValue}><i className="ms-Icon ms-Icon--FabricFolderFill" title="FabricFolderFill" aria-hidden="true" style={{fontSize: "15px",paddingRight: "5px",color: "rgb(255,185,0)"}}></i><span>{item.LinkFilename}</span></a></>;
        }
        return <><a data-intercenption="off" target="_blank" id="anchorStyleFile" href={hrefValue}><i className="ms-Icon ms-Icon--DocumentSet" title="FabricFolderFill" aria-hidden="true" style={{fontSize: "15px",paddingRight: "5px",color: "rgb(214,85,50)"}}></i><span>{item.LinkFilename}</span></a></>;
      }},
      { key: 'folderName', name: 'Channel Name', fieldName: 'FileDirRef', minWidth: 100, maxWidth: 100, isResizable: true ,
        onRender: (item: any) => {
          return <span style={{fontWeight: "bold", fontSize: "14px"}}>{item.FileDirRef}</span>
        }
      },
      { key: 'Author', name: 'Created By', fieldName: 'Author', minWidth: 100, maxWidth: 100, isResizable: true ,
      onRender: (item: any) => {
        return <span style={{fontWeight: "italic", fontSize: "12px"}}>{item.Author.Title}</span>
      }},
      { key: 'Modified', name: 'Modified By', fieldName: 'Editor', minWidth: 100, maxWidth: 100, isResizable: true ,
      onRender: (item: any) => {
        return <span style={{fontWeight: "italic", fontSize: "12px"}}>{item.Editor.Title}</span>
      }},
      { key: 'folderPath', name: 'File/Folder Path', fieldName: 'FileRef', minWidth: 100, maxWidth: 300, isResizable: true }
    ];

    getSP();
    this.ListServeInstance = new SPInstance();
}

  public componentDidMount(): void {
    var serverurl = this.props.ctx.pageContext.web.serverRelativeUrl;
    this.ListServeInstance.getListItems("Documents").then(values =>{
      _.forEach(values, function(item){
        var folderName  = item.FileDirRef.replace(serverurl,"").split('/')
        item.FileDirRef = folderName[2]  === undefined ? "" : folderName[2];
      });
      this.setState({
        items: values,
        filteredItems: values
      })
  }).catch(e => {
    console.log(e);
  }); 
}

private _onRenderRow: IDetailsListProps['onRenderRow'] = props => {
  const customStyles: Partial<IDetailsRowStyles> = {};
  if (props) {
    if (props.itemIndex % 2 !== 0) {
      // Every other row renders with a different background color
      customStyles.root = { backgroundColor: "rgb(199 233 217)" };
    }

    return <DetailsRow {...props} styles={customStyles} />;
  }
  return null;
};

private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
//   var filteredItems = (this.state.items, function(o) { 
//     return o.Title == 'john'; 
//  });
  this.setState({
    filteredItems: text ? this.state.items.filter((i: { LinkFilename: string; }) => i.LinkFilename.toLowerCase().indexOf(text) > -1) : this.state.items,
  });
};

  public render(): React.ReactElement<ISearchDocumentsProps> {
    const {
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.searchDocuments} ${hasTeamsContext ? styles.teams : ''}`}>
        <img className='headerImg' alt="" src={isDarkTheme ? require('../assets/search.jpg') : require('../assets/search.jpg')}/>
        
        <div style={{border: "2px solid darkgrey",margin: "1% 10% 1% 10%"}}></div>
        <div className="bxShadow">
        <div style={{fontSize: "20px",fontWeight: "bold",textAlign: "center",padding: "10px", background:"#e0dfdf", color:"#a60505"}}>Search Folders or Files across all public channels</div>
        <div className='ms-Grid'>
        <div className={`ms-Grid-row ${styles.row}`}>
          <div className='ms-Grid-col ms-sm2 ms-md2'>
          <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/People.png') : require('../assets/People.png')} className={styles.welcomeImage} />
          <h3 style={{color: "#a60505"}}>Welcome, {escape(userDisplayName)}!</h3>
          <div className='howDoes'>How does this work:</div>
          <div className='points'>- Search any of your files or folder across the public channels.</div>
          <div className='points'>- Navigate to the file to read or edit.</div>
        </div>
          </div>
          <div className='ms-Grid-col ms-sm10 ms-md10'>
          <TextField label="Filter by Title:" onChange={this._onChangeText} />
            <DetailsList items={this.state.filteredItems} columns={this._columns} onRenderRow={this._onRenderRow}/>
          </div>
        </div>
        </div>
        </div>
      </section>
    );
  }
}
