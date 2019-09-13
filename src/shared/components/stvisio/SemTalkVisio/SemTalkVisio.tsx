import * as React from 'react';
import styles from './SemTalkVisio.module.scss';
// import { SemTalkBreadCrumbs } from '../../stbreadcrumbs/SemTalkBreadCrumbs';
// import { SemTalkCommandBar } from '../../stcommandbar/SemTalkCommandBar';
// import { SemTalkPivot } from '../../stpivot/SemTalkPivot';
import { autobind, } from 'office-ui-fabric-react';
// import {
//   FindModelByID, ModelTable, FindDiagramByNameID,
//   DiagramTable, FindModelByName
// } from '../../../../shared/semtalkportal/dbase';
// import { SetContext } from '../../../../shared/semtalkportal/restinterface';
//import { Clear } from '../../../../shared/semtalkportal/stclear';
// import { getHostPageInfoListener, setModel } from '../../../../shared/semtalkportal/stglobal';
//import { setModel } from '../../../../shared/semtalkportal/stglobal';
//import { setURL, setDiagram } from '../../../../shared/semtalkportal/stglobal';
// import { DoArgs } from '../../../../shared/semtalkportal/stglobal';
//import { addCallBack, removeCallBack, setStay } from '../../../../shared/semtalkportal/stglobal';
import { MSGraphClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

import { VisioService } from "../../../../shared/services";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Guid } from "guid-typescript";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

/* import {
  FindObject
} from '../../../../shared/semtalkportal/jbase';
 */
initializeIcons(undefined, { disableWarnings: true });

export interface ISemTalkVisioProps {
  visioService: VisioService;
  context: WebPartContext;
  documentUrl: string;
  filter: string;
  site: string;
  document: string;
  width: string;
  height: string;
  showDetails: boolean;
  showProps: boolean;
  showCon: boolean;
  showtype?: boolean;
  shownodes?: boolean;
  showproperties?: boolean;
  shownav?: boolean;
  hidebpmn?: string[];
  hidesimulation?: string[];
  showPropsGrouped: boolean;
  showLinks: boolean;
  showDiagram: boolean;
  showDocument: boolean;
  showAttachment: boolean;
  showTeams: boolean;
  showBreadCrumbs: boolean;
  showCommandBar: boolean;
  showListItems: boolean;
  showListDocuments: boolean;
  editDocumentList: boolean;
  editList: boolean;
  listitemssite?: string;
  listitemslist?: string;
  listitemsquery?: string;
  listitemscolumns?: string[];
  listdocumentslibrary?: string;
  listdocumentssite?: string;
  listdocumentscolumns?: string[];
  listdocumentsquery?: string;
  islist: boolean;
  gotonodes: boolean;
  goodlist: string[];
  badlist: string[];
  objroot: string;
  navroot: string;
  views: string;
  nrOfItems?: number | undefined;
  displayMode?: DisplayMode | undefined;
  graphClient?: MSGraphClient;
  updateProperty?: (value: string) => void;
  usegraph: boolean;
  defaulttopic?: string;
  portal?: string;
  teamid?: string;
  planid?: string;
  showBot: boolean;
  botsecret: string;
  showWiki?: boolean;
  wikilist?: string | undefined;
  wikisite?: string | undefined;
  wikieditList?: boolean;
  isSingleDocument: boolean;
  navVisable?: boolean;
  searchVisable?: boolean;
  propsVisable?: boolean;
  diagVisable?: boolean;
  hieVisable?: boolean;
  teaVisable?: boolean;
  docVisable?: boolean;
  procVisable?: boolean;
  lisVisable?: boolean;
  wikiVisable?: boolean;
  detVisable?: boolean;
  linVisable?: boolean;
  propGVisable?: boolean;
  useVisable?: boolean;
  trendVisable?: boolean;
  whoVisable?: boolean;
  conVisable?: boolean;
  repVisable?: boolean;
  roleVisable?: boolean;
  docInfoVisable?: boolean;
  botVisable?: boolean;
  objIsList?: boolean;
  objIscombo?: boolean;
  objGoodlist?: string;
  objBadlist?: string;
}
export interface ISemTalkVisioState {
  // selectedShape: Visio.Shape | null;
  selectedPage: Visio.Page | null;
  selectedDocument: Visio.Document | null;
  modelid: number;
  loading: boolean;
}
export class SemTalkVisio extends React.Component<ISemTalkVisioProps, ISemTalkVisioState> {
  private xmlurl: string = "";
  // private _id: string;
  // private _height: string = "";
  private _site: string;
  private currentPage: string = "";
  private loadedforshape: string = "";
  private loadedforpage: string = "";
 // private navigate: boolean = true;
  public callback: Guid;

  constructor(props: ISemTalkVisioProps) {
    super(props);
    console.debug("SemTalkVisio");

    // set delegate functions that will be used to pass the values from the Visio service to the component
    if (this.props.visioService != undefined) {
      this.props.visioService.onSelectionChanged = this._onSelectionChanged;
      this.props.visioService.onPageLoadComplete = this._onPageLoadComplete;
      this.props.visioService.onDocumentLoadComplete = this._onDocumentLoadComplete;
      this.props.visioService.getAllShapes = this._getAllShapes;
    }
    this.callback = Guid.create();
    this.state = {
      // selectedShape: null,
      selectedPage: null,
      selectedDocument: null,
      modelid: -1,
      loading: true
    };
    /*    FindObject("Class","Event").then((e)=> {
         console.log(e);
       }); */

  }
  public render(): React.ReactElement<ISemTalkVisioProps> {
    // let fil = {
    //   filtermodel: "Model",
    //   filterpage: "Page",
    //   filterobject: "Object",
    //   filtershape: "ShapeID"
    // };
    let divHostStyle = {
      height: this.props.height,
      width: this.props.width
    };
    return (
      <div>
        <div>
         </div> <div>
          {/* {this.props.showCommandBar &&
          //     <SemTalkCommandBar
          //       context={this.props.context}
          //       filter={this.props.filter}
          //       islist={this.props.islist}
          //       goodlist={this.props.goodlist}
          //       badlist={this.props.badlist}
          //       objroot={this.props.objroot}
          //       navroot={this.props.navroot}
          //       site={this.props.listitemssite}
          //       listTitle={this.props.listitemslist}
          //       query={this.props.listitemsquery}
          //       columns={this.props.listitemscolumns}
          //       dsite={this.props.listdocumentssite}
          //       dTitle={this.props.listdocumentslibrary}
          //       dquery={this.props.listdocumentsquery}
          //       dcolumns={this.props.listdocumentscolumns}
          //       editDocumentList={this.props.editDocumentList}
          //       editList={this.props.editList}
          //       listFilter={fil}
          //       addnew={true}
          //       nrOfItems={this.props.nrOfItems}
          //       displayMode={this.props.displayMode}
          //       updateProperty={this.props.updateProperty}
          //       defaulttopic={this.props.defaulttopic}
          //       graphClient={this.props.graphClient}
          //       usegraph={this.props.usegraph}
          //       teamid={this.props.teamid}
          //       planid={this.props.planid}
          //       portal={this.props.portal}
          //       botsecret={this.props.botsecret}
          //       views={this.props.views}
          //       wikilist={this.props.wikilist}
          //       wikisite={this.props.wikisite}
          //       wikieditList={this.props.wikieditList}
          //       navVisable={this.props.navVisable}
          //       searchVisable={this.props.searchVisable}
          //       propsVisable={this.props.propsVisable}
          //       diagVisable={this.props.diagVisable}
          //       hieVisable={this.props.hieVisable}
          //       teaVisable={this.props.teaVisable}
          //       docVisable={this.props.docVisable}
          //       procVisable={this.props.procVisable}
          //       lisVisable={this.props.lisVisable}
          //       wikiVisable={this.props.wikiVisable}
          //       detVisable={this.props.detVisable}
          //       linVisable={this.props.linVisable}
          //       propGVisable={this.props.propGVisable}
          //       useVisable={this.props.useVisable}
          //       trendVisable={this.props.trendVisable}
          //       whoVisable={this.props.whoVisable}
          //       conVisable={this.props.conVisable}
          //       repVisable={this.props.repVisable}
          //       roleVisable={this.props.roleVisable}
          //       docInfoVisable={this.props.docInfoVisable}
          //       botVisable={this.props.botVisable}
          //       objIsList={this.props.objIsList}
          //       objIscombo={this.props.objIscombo}
          //       objGoodlist={this.props.objGoodlist}
          //       objBadlist={this.props.objBadlist}

          //     />
          //   }
          //   {this.props.showBreadCrumbs &&
          //     <SemTalkBreadCrumbs context={this.props.context} filter={this.props.filter} />
          //   }
          // </div>
          */}
          <div className={styles.semTalkVisio}>
            <div id='iframeHost' className={styles.iframeHost} style={divHostStyle}></div>
            {/* {(this.props.showDetails || this.props.showDocument || this.props.showProps || this.props.showCon || this.props.showPropsGrouped
              || this.props.showLinks || this.props.showDiagram || this.props.showAttachment || this.props.showListItems || this.props.showListDocuments) &&
              <div className={styles.detailsPanel} >
                <SemTalkPivot
                  context={this.props.context}
                  filter={this.props.filter}
                  gotonodes={this.props.gotonodes}
                  showDetails={this.props.showDetails}
                  showDocument={this.props.showDocument}
                  showProps={this.props.showProps}
                  showContext={this.props.showCon}
                  showtype={this.props.showtype}
                  shownodes={this.props.shownodes}
                  showproperties={this.props.showproperties}
                  shownav={this.props.shownav}
                  hidebpmn={this.props.hidebpmn}
                  hidesimulation={this.props.hidesimulation}
                  showPropsGrouped={this.props.showPropsGrouped}
                  showLinks={this.props.showLinks}
                  showDiagram={this.props.showDiagram}
                  showAttachment={this.props.showAttachment}
                  showTeams={this.props.showTeams}
                  showListItems={this.props.showListItems}
                  showListDocuments={this.props.showListDocuments}
                  listitemssite={this.props.listitemssite}
                  listitemslist={this.props.listitemslist}
                  listitemsquery={this.props.listitemsquery}
                  listitemscolumns={this.props.listitemscolumns}
                  editList={this.props.editList}
                  listdocumentslibrary={this.props.listdocumentslibrary}
                  listdocumentssite={this.props.listdocumentssite}
                  listdocumentsquery={this.props.listdocumentsquery}
                  listdocumentscolumns={this.props.listdocumentscolumns}
                  editDocumentList={this.props.editDocumentList}
                  views={this.props.views}
                  goodlist={this.props.goodlist}
                  badlist={this.props.badlist}
                  // goto_context={getContext()}
                  listFilter={fil}
                  addnew={true}
                  graphClient={this.props.graphClient}
                  usegraph={this.props.usegraph}
                  // goto_context={getContext()}
                  teamid={this.props.teamid}
                  portal={this.props.portal}
                  planid={this.props.planid}
                  showBot={this.props.showBot}
                  botsecret={this.props.botsecret}
                  showWiki={this.props.showWiki}
                  wikisite={this.props.wikisite}
                  wikilist={this.props.wikilist}
                  wikieditList={this.props.wikieditList}
                />
              </div>
            } */}
          </div>
        </div>
      </div>
    );
  }
  //  private mounted: boolean = false;
  public componentDidMount() {
    if (this.props.context && this.props.filter) {
    //  SetContext(this.props.context, this.props.filter);
    }
   // setStay(this.props.isSingleDocument);
    if (this.props.documentUrl) {
      // let mid: number = -1;
      // this.xmlurl = this.props.documentUrl;
      this._site = this.props.site;
      //   // this._id = "B896486FB-88F7-4CF6-8547-4DDC9FC7E638";
      this.props.visioService.load(this.props.documentUrl, this.props.width, this.props.height)
        .then(() => {
          let mname = this._site + this.props.document.replace("/VSDX/", "/XML/").replace(".vsdx", ".xml");
          this.xmlurl = mname;
          // return FindModelByName(mname);
        });
        // .then((m: ModelTable) => {
        //   mid = m.ID;
        //   setURL(mid);
        //   this.setState({ modelid: mid });
        //   setModel(mid);
        //   return this.props.visioService.activePage();
        // })
        // .then((pg: Visio.Page) => {
        //   return FindDiagramByNameID(pg.name, mid);
        // })
        // .then((d: DiagramTable) => {
        //   setDiagram(d.ID);
        //   this._id = mid.toString();
        //   console.log("Startpage: ", d.ObjectCaption);
        // });
    }
   // addCallBack(this, (e: CustomEvent) => this.eventListener(e));
   // DoArgs();
  }
  public componentWillUnmount() {
   // removeCallBack(this, (e: CustomEvent) => this.eventListener(e));
  }
  public async componentDidUpdate(prevProps: ISemTalkVisioProps) {
  //  setStay(this.props.isSingleDocument);
  //  setModel(this.state.modelid);
    if (this.props.documentUrl && this.props.documentUrl !== prevProps.documentUrl) {
      this.props.visioService.load(this.props.documentUrl, this.props.width, this.props.height);
    }
  }
  /* private isKey(evt: any, key: string): boolean {
    switch (key) {
      case "ctrl":
        return evt.ctrlKey == true;
      case "shift":
        return evt.shiftKey == true;
      case "":
        return true;
    }
    return false;
  } */
  /**
   * method executed after a on selection change event is triggered
   * @param selectedShape the shape selected by the user on the Visio diagram
   */
  private _onSelectionChanged = (selectedShape: Visio.Shape): void => {
    if (this.state.loading) {
      return;
    }
/*     if (selectedShape != null && this.state.selectedPage != null) {
      var keyInfo: any = window.event;
      if (keyInfo.shiftKey) {
      }
      var md: any = {};
      md.modelname = this.xmlurl;

      if (selectedShape.hyperlinks != null && selectedShape.hyperlinks.items != null &&
        selectedShape.hyperlinks.items.length == 1 && this.navigate) {
        this.navigate = true;
        let hyp: any = selectedShape.hyperlinks.items[0];
        if (hyp.address.length > 0) {

          //   window.open(hyp.address, "_blank");
          return;
        } else {
          md.type = "GotoPage";
          md.pagename = hyp.subAddress;
          // md.refine = this.isKey(evt, this.props.svgrefine);
          // md.prop = this.isKey(evt, this.props.svgprop);
          // md.hyperlink = this.isKey(evt, this.props.svghyperlink);
          getHostPageInfoListener(JSON.stringify(md));
          return;
        }
      }
      md.type = "shapeSelectionChanged";
      md.modelname = this.xmlurl;
      md.pagename = this.state.selectedPage.name;
      md.shapeID = "Sheet." + selectedShape.id;  //args.shapeNames[0];
      if (this.navigate) {
        md.refine=true;
        md.prop=false;
      }
      this.navigate = true;

      // md.refine=this.isKey(evt, this.props.svgrefine);
      // md.prop=this.isKey(evt, this.props.svgrefine);
      // md.hyperlink=this.isKey(evt, this.props.svghyperlink);
      getHostPageInfoListener(JSON.stringify(md));
    } else {
    } */

  }
  private _onPageLoadComplete = (selectedPage: Visio.Page): void => {

    console.log("Page load complete: ", selectedPage.name);
    this.setState({
      selectedPage: selectedPage,
      loading: false
    });
 /*    this.currentPage = selectedPage.name;
    var md: any = {};
    md.type = "GotoPage";
    md.modelname = this.xmlurl;
    md.pagename = selectedPage.name;
    getHostPageInfoListener(JSON.stringify(md)); */

  }
  private _onDocumentLoadComplete = (selectedDocument: Visio.Document): void => {

    console.log("Document load complete: ", selectedDocument);
    this.setState({
      selectedDocument: selectedDocument
    });
   // SetContext(this.props.context, this.props.filter);
    // AllModels().then((m: any) => {
    //   console.log("Model: ", m);
    // });
    if (this.loadedforpage != "") {
      this.props.visioService.selectPage(this.loadedforpage);
      this.loadedforpage = "";
    } else {
      this.props.visioService.activePage().then((pg: Visio.Page) => {
        this._onPageLoadComplete(pg);
      });
    }

  }

  /**
 * method executed after the collection of shapes is retrieved - after Visio diagram page load
 * @param shapes the collection of shapes on the Visio diagram
 */

  private _getAllShapes = (shapes: Visio.Shape[]): void => {
     if (this.loadedforshape != "") {
      this.gotoShape(this.loadedforshape);
      this.loadedforshape = "";
    }
    // console.log("Models: ", models);

  }
 @autobind
  public async handleEvent(m: any): Promise<void> {
    var mstr = JSON.stringify(m);
    this.eventListener({ data: mstr });
  }
  @autobind
  private eventListener(e: any): void {
    let md: any;
    try {
      //  console.log(e.data);
      md = JSON.parse(e.data);
    }
    catch (error) {
      // console.log("Could not parse the message response.");
      return;
    }
    var mtype = md.type;
    switch (mtype) {
      case "gotoDocument": {
        let pagename = md.pagename;
        let modelname = md.modelname;
        let d = pagename;
        if (d.indexOf("#") > 0) {
          d = d.substring(d.indexOf("#") + 1);
        }
        this.loadedforshape = "";

        if (this.xmlurl != modelname) {
          this.xmlurl = modelname;
          // FindModelByID(md.modelid).then((mod: ModelTable) => {
          //   let uid = mod.UniqueID;
          //   uid = uid.replace("{", "");
          //   uid = uid.replace("}", "");
          //   if (uid != this._id) {
          //     this.loadedforpage = d;
          //     this.gotoVisioOnlineDocument(this._site, uid, this._height);
          //     return;
          //   }
          // });
        } else {
          if (this.state.selectedPage == null || this.state.selectedPage.name != d) {
            this.props.visioService.selectPage(d);
          }
        }
      }
        break;
      case "gotoObject":
        break;
      case "gotoNode": {
        let txt = "";
        let val = "";
        this.gotoNode(md.modelid, md.modelname, md.pagename, txt, md.shapeid, val);
      }
        break;
      case "gotoShape": {
        let txt = "";
        let val = "";
        this.gotoNode(md.modelid, md.modelname, md.pagename, txt, md.shapeid, val);
      }
        break;
    }
    if (md.message == "getHostPageInfo") {
    }
    if (md.message == "getHostSemTalkHomePage") {
    }
  }

  private gotoNode(modelid: number, md: string, pg: string, txt: string, sh: string, val: string) {
    try {
      let d = pg;
      if (d.indexOf("#") > 0) {
        d = d.substring(d.indexOf("#") + 1);
      }
      this.loadedforshape = sh;
      this.loadedforpage = "";

      if (md != this.xmlurl) {
        this.xmlurl = md;
        // FindModelByID(modelid).then((mod: ModelTable) => {
        //   let uid = mod.UniqueID;
        //   uid = uid.replace("{", "");
        //   uid = uid.replace("}", "");
        //   if (uid != this._id) {
        //     this.loadedforpage = d;
        //     this.gotoVisioOnlineDocument(this._site, uid, this._height);
        //     return;
        //   }
        // });
      }
      if (this.currentPage != d) {
        this.props.visioService.selectPage(d);
        // .then(() => {
        //   this.gotoShape(sh);
        // });
      } else {
        this.gotoShape(sh);
      }
      // }
    }
    catch (e) {
      alert(pg + ': ' + txt);
    }
  }
  private gotoShape(sh: string) {
   // this.navigate = false;
    this.props.visioService.selectShapeByID(sh);
  }
/*   private gotoVisioOnlineDocument(site: string, id: string, h: string): void {
    this._site = site;
    this._id = id;
    this._height = h;
    this.setState({
      loading: true
    });
    Clear("iframeHost");
    let d = site + '/_layouts/15/Doc.aspx?sourcedoc=%7B' + id + '%7D&action=embedview';
    this.props.visioService.load(d, this.props.width, this.props.height)
      .then(() => {
        if (this.loadedforpage != "") {
          this.props.visioService.selectPage(this.loadedforpage);
        }
      });
  } */
}
