

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneHorizontalRule

} from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';

import 'officejs';
// import * as strings from 'SemTalkVisioWebPartStrings';
import * as strings from 'SemTalkStrings';
import { SemTalkVisio, ISemTalkVisioProps } from '../../shared/components/stvisio/SemTalkVisio';
import { VisioService } from "../../shared/services";
//import { stconfig } from '../../shared/semtalkportal/stconfig';

export interface ISemTalkVisioWebPartProps {
  documentUrl: string;
  service: string;
  filter: string;
  site: string;
  document: string;
  width: string;
  height: string;
  showDetails: boolean;
  showProps: boolean;
  showContext: boolean;
  showtype: boolean;
  shownodes: boolean;
  showproperties: boolean;
  shownav: boolean;
  hidebpmn: string;
  hidesimulation: string;
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
  listitemssite: string;
  listitemslist: string;
  listitemsquery: string;
  listitemscolumns: string;
  listdocumentslibrary: string;
  listdocumentssite: string;
  listdocumentsquery: string;
  listdocumentscolumns: string;
  editDocumentList: boolean;
  usegraph: boolean;
  editList: boolean;
  islist: boolean;
  isSingleDocument: boolean;
  goodlist: string;
  badlist: string;
  objroot: string;
  navroot: string;
  views: string;
  gotonodes: boolean;
  title: string;
  nrOfItems: number;
  defaulttopic: string;
  teamid: string;
  planid: string;
  portal: string;
  showBot: boolean;
  botsecret: string;
  showWiki?: boolean;
  wikilist: string;
  wikisite: string;
  wikieditList: boolean;
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

export default class SemTalkVisioWebPart extends BaseClientSideWebPart<ISemTalkVisioWebPartProps> {
  private graphClient: MSGraphClient;
  // private _teamsContext: microsoftTeams.Context;

  // This variable has been added
  private _visioService: VisioService;
  public onInit(): Promise<void> {
    // let retVal: Promise<any> = Promise.resolve();
    // if (this.context.microsoftTeams) {
    //   retVal = new Promise((resolve, _reject) => {
    //     this.context.microsoftTeams.getContext(context => {
    //       this._teamsContext = context;
    //       resolve();
    //     });
    //   });
    // }
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      console.log("Mock data service not implemented yet");
    } else {
      this._visioService = new VisioService(this.context);
    }
    if (this.properties.service) {
  //    stconfig._service = this.properties.service;
    }
    if (this.properties.filter) {
    //  stconfig._filter = this.properties.filter;
    }
    if (this.properties.listitemssite == undefined) {
      console.debug("listitemssite undefined");
    }

    return super.onInit();


  }

  public render(): void {
    console.log("Render Visio Webpart");
  //  this.properties.views = JSON.stringify(views);
    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        this.graphClient = client;
      }).then(() => {

        let gl: string[] = [];
        if (this.properties.goodlist && this.properties.goodlist.length > 0) {
          // console.debug("Goodlist: ",this.properties.goodlist);
          gl = this.properties.goodlist.split(", ");
        }
        let bl: string[] = [];
        if (this.properties.badlist && this.properties.badlist.length > 0) {
          // console.debug("Badlist: ",this.properties.badlist);
          bl = this.properties.badlist.split(", ");
        }
        let cl: string[] = [];
        if (this.properties.listitemscolumns && this.properties.listitemscolumns.length > 0) {
          cl = this.properties.listitemscolumns.split(", ");
        }
        let bpmnl: string[] = [];
        if (this.properties.hidebpmn && this.properties.hidebpmn.length > 0) {
         // console.debug("BPMNlist: ", this.properties.hidebpmn);
          bpmnl = this.properties.hidebpmn.split(", ");
        }
        let siml: string[] = [];
        if (this.properties.hidesimulation && this.properties.hidesimulation.length > 0) {
        //  console.debug("Simlist: ", this.properties.hidesimulation);
          siml = this.properties.hidesimulation.split(", ");
        }
        const element: React.ReactElement<ISemTalkVisioProps> = React.createElement(
          SemTalkVisio,
          {
            visioService: this._visioService,
            context: this.context,
            documentUrl: this.properties.documentUrl,
            site: this.properties.site,
            filter: this.properties.filter,
            document: this.properties.document,
            width: this.properties.width,
            height: this.properties.height,
            showDetails: this.properties.showDetails,
            showProps: this.properties.showProps,
            showCon: this.properties.showContext,
            showPropsGrouped: this.properties.showPropsGrouped,
            showLinks: this.properties.showLinks,
            showDiagram: this.properties.showDiagram,
            showDocument: this.properties.showDocument,
            showAttachment: this.properties.showAttachment,
            showTeams: this.properties.showTeams,
            showBreadCrumbs: this.properties.showBreadCrumbs,
            showListDocuments: this.properties.showListDocuments,
            showCommandBar: this.properties.showCommandBar,
            gotonodes: this.properties.gotonodes,
            islist: this.properties.islist,
            goodlist: gl,
            badlist: bl,
            showtype: this.properties.showtype,
            shownodes: this.properties.shownodes,
            showproperties: this.properties.showproperties,
            shownav: this.properties.shownav,
            hidebpmn: bpmnl,
            hidesimulation: siml,
            views: this.properties.views,
            objroot: this.properties.objroot,
            navroot: this.properties.navroot,
            showListItems: this.properties.showListItems,
            listitemssite: this.properties.listitemssite,
            listitemslist: this.properties.listitemslist,
            listitemsquery: this.properties.listitemsquery,
            listitemscolumns: cl,
            listdocumentslibrary: this.properties.listdocumentslibrary,
            listdocumentssite: this.properties.listdocumentssite,
            editDocumentList: this.properties.editDocumentList,
            editList: this.properties.editList,
            graphClient: this.graphClient,
            usegraph: this.properties.usegraph,
            nrOfItems: this.properties.nrOfItems,
            displayMode: this.displayMode,
            updateProperty: (value: string) => {
              this.properties.title = value;
            },
            defaulttopic: this.properties.defaulttopic,
            teamid: this.properties.teamid,
            planid: this.properties.planid,
            portal: this.properties.portal,
            showBot: this.properties.showBot,
            botsecret: this.properties.botsecret,
            showWiki: this.properties.showWiki,
            wikilist: this.properties.wikilist,
            wikisite: this.properties.wikisite,
            wikieditList: this.properties.wikieditList,
            isSingleDocument: this.properties.isSingleDocument,
            navVisable: this.properties.navVisable,
            searchVisable: this.properties.searchVisable,
            propsVisable: this.properties.propsVisable,
            diagVisable: this.properties.diagVisable,
            hieVisable: this.properties.hieVisable,
            teaVisable: this.properties.teaVisable,
            docVisable: this.properties.docVisable,
            procVisable: this.properties.procVisable,
            lisVisable: this.properties.lisVisable,
            wikiVisable: this.properties.wikiVisable,
            detVisable: this.properties.detVisable,
            linVisable: this.properties.linVisable,
            propGVisable: this.properties.propGVisable,
            useVisable: this.properties.useVisable,
            trendVisable: this.properties.trendVisable,
            whoVisable: this.properties.whoVisable,
            conVisable: this.properties.conVisable,
            repVisable: this.properties.repVisable,
            roleVisable: this.properties.roleVisable,
            docInfoVisable: this.properties.docInfoVisable,
            botVisable: this.properties.botVisable,
            objIsList: this.properties.objIsList,
            objIscombo: this.properties.objIscombo,
            objGoodlist: this.properties.objGoodlist,
            objBadlist: this.properties.objBadlist,
          }
        );
        ReactDom.render(element, this.domElement);
      });
    // }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.UIGroupName,
              groupFields: [
                PropertyPaneToggle('showProps', {
                  label: strings.PropsLabel
                }),
                // PropertyPaneToggle('showtype', {
                //   label: strings.TypeLabel
                // }),
                PropertyPaneTextField('width', {
                  label: "Width:",
                }),
                PropertyPaneTextField('height', {
                  label: "Height:",
                }),
                PropertyPaneToggle('shownodes', {
                  label: strings.NodesLabel
                }),
                PropertyPaneToggle('showproperties', {
                  label: strings.AssocLabel
                }),
                PropertyPaneToggle('shownav', {
                  label: strings.NavLabel
                }),
                PropertyPaneTextField('hidebpmn', {
                  label: strings.BPMNLabel
                }),
                PropertyPaneTextField('hidesimulation', {
                  label: strings.SimLabel
                }),
                PropertyPaneToggle('gotonodes', {
                  label: strings.GotoLabel
                }),
                PropertyPaneTextField('goodlist', {
                  label: strings.GoodLabel
                }),
                PropertyPaneTextField('badlist', {
                  label: strings.BadLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showContext', {
                  label: strings.ContextLabel
                }),
               PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showPropsGrouped', {
                  label: strings.PropsGroupedLabel
                }),
                PropertyPaneTextField('views', {
                  label: strings.GrpPropViews,
                  multiline: true
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showDetails', {
                  label: strings.DetailsLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showDiagram', {
                  label: strings.DiagramLabel
                }),
                PropertyPaneToggle('showDocument', {
                  label: strings.DocInfoLabel
                }),
                PropertyPaneToggle('showLinks', {
                  label: strings.LinksLabel
                }),
                PropertyPaneToggle('showAttachment', {
                  label: strings.AttachmentLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showTeams', {
                  label: strings.TeamsLabel
                }),
                PropertyPaneTextField('teamid', {
                  label: strings.TeamsID,
                }),
                PropertyPaneTextField('planid', {
                  label: strings.PlannerID,
                }),
                PropertyPaneTextField('portal', {
                  label: strings.Portal,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showBot', {
                  label: strings.BotLabel
                }),
                PropertyPaneTextField('botsecret', {
                  label: strings.BotSecret,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showBreadCrumbs', {
                  label: strings.BreadCrumbsLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showCommandBar', {
                  label: strings.CommandBarLabel
                }),
                PropertyPaneToggle('islist', {
                  label: strings.IsListLabel
                }),
                PropertyPaneTextField('objroot', {
                  label: strings.ObjRootLabel
                }),
                PropertyPaneTextField('navroot', {
                  label: strings.NavRootLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showListItems', {
                  label: strings.ListItemsLabel
                }),
                PropertyPaneTextField('listitemssite', {
                  label: strings.SiteLabel,
                }),
                PropertyPaneTextField('listitemslist', {
                  label: strings.List,
                }),
                PropertyPaneToggle('editList', {
                  label: strings.Edit,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showListDocuments', {
                  label: strings.ListDocumentsLabel
                }),
                PropertyPaneTextField('listdocumentslibrary', {
                  label: strings.LibraryLabel,
                }),
                PropertyPaneTextField('listdocumentssite', {
                  label: strings.LibrarySiteLabel,
                }),
                PropertyPaneToggle('editDocumentList', {
                  label: strings.Edit,
                }),
                PropertyPaneToggle('usegraph', {
                  label: strings.GraphProps,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('showWiki', {
                  label: strings.WikiLabel
                }),
                PropertyPaneTextField('wikilist', {
                  label: strings.WikiListLabel
                }),
                PropertyPaneTextField('wikisite', {
                  label: strings.WikiSiteLabel
                }),
                PropertyPaneToggle('wikieditList', {
                  label: strings.Edit,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('isSingleDocument', {
                  label: strings.SingleDocument,
                }),
              ]
            },
            {
              groupName: strings.Hierarchy,
              groupFields: [
                PropertyPaneTextField('objroot', {
                  label: strings.ObjRootLabel
                }),
                PropertyPaneToggle('objIslist', {
                  label: strings.IsListLabel
                }),
                PropertyPaneToggle('objIscombo', {
                  label: strings.IsComboLabel
                }),
                PropertyPaneTextField('objGoodlist', {
                  label: strings.GoodLabel
                }),
                PropertyPaneTextField('objBadlist', {
                  label: strings.BadLabel
                }),
              ]
            },
            {
              groupName: strings.CommandBarEntries,
              groupFields: [
                PropertyPaneToggle('navVisable', {
                  label: strings.Navigation
                }),
                PropertyPaneToggle('searchVisable', {
                  label: strings.Search
                }),
                PropertyPaneToggle('propsVisable', {
                  label: strings.PropsLabel
                }),
                PropertyPaneToggle('diagVisable', {
                  label: strings.Diagram
                }),
                PropertyPaneToggle('hieVisable', {
                  label: strings.Hierarchy
                }),
                PropertyPaneToggle('teaVisable', {
                  label: strings.Planner
                }),
                PropertyPaneToggle('docVisable', {
                  label: strings.Documents
                }),
                PropertyPaneToggle('procVisable', {
                  label: strings.Process
                }),
                PropertyPaneToggle('lisVisable', {
                  label: strings.List
                }),
                PropertyPaneToggle('wikiVisable', {
                  label: strings.WikiLabel
                }),
                PropertyPaneToggle('detVisable', {
                  label: strings.DetailsLabel
                }),
                PropertyPaneToggle('linVisable', {
                  label: strings.LinksLabel
                }),
                PropertyPaneToggle('conVisable', {
                  label: strings.ContextLabel
                }),
                PropertyPaneToggle('propGVisable', {
                  label: strings.PropsGroupedLabel
                }),
                PropertyPaneToggle('useVisable', {
                  label: strings.UsedDocuments
                }),
                PropertyPaneToggle('trendVisable', {
                  label: strings.TrendingDocuments
                }),
                PropertyPaneToggle('whoVisable', {
                  label: strings.People
                }),
                PropertyPaneToggle('repVisable', {
                  label: strings.Reports
                }),
                PropertyPaneToggle('roleVisable', {
                  label: strings.RoleAssignment
                }),
                PropertyPaneToggle('docInfoVisable', {
                  label: strings.DocInfoLabel
                }),
                PropertyPaneToggle('botVisable', {
                  label: strings.Bot
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('documentUrl', {
                  label: strings.DocumentUrlLabel
                }),
                PropertyPaneTextField('site', {
                  label: strings.SiteLabel
                }),
                PropertyPaneTextField('document', {
                  label: strings.DocumentLabel
                })
              ]
            },
            {
              groupName: strings.BackendGroupName,
              groupFields: [
                PropertyPaneTextField('service', {
                  label: strings.ServiceLabel
                }),
                PropertyPaneTextField('filter', {
                  label: strings.FilterLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
