declare interface ISemTalkStrings {

  TitleLabel: string;
  DescriptionLabel: string;
  WidthLabel: string;
  HeightLabel: string;
  MinimizeLabel: string;
  HelpLabel: string;
  SPListnameLabel: string;
  SPURLLabel: string;
  WebPartGroupName: string;
  Query: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DocumentUrlLabel: string;
  SiteLabel: string;
  DocumentLabel: string;
  BackendGroupName: string;
  ServiceLabel: string;
  FilterLabel: string;
  UIGroupName: string;
  DetailsLabel: string;
  PropsLabel: string;
  ContextLabel: string;
  PropsGroupedLabel: string;
  GotoLabel: string;
  LinksLabel: string;
  DiagramLabel: string;
  DocInfoLabel: string;
  AttachmentLabel: string;
  DataObjectAttachmentLabel: string;
  TeamsLabel: string;
  BotLabel: string;
  ListItemsLabel: string;
  WikiLabel: string;
  WikiListLabel: string;
  WikiSiteLabel: string;
  BreadCrumbsLabel: string;
  CommandBarLabel: string;
  IsListLabel: string;
  IsComboLabel: string;
  GoodLabel: string;
  BadLabel: string;
  ObjRootLabel: string;
  NavRootLabel: string;
  GrpPropViews: string;
  ListDocumentsLabel: string;

//Roles
  RootLabel: string;

 //Reports
   PageSizeLabel: string;
  ReportNameLabel: string;
  IsFilterLabel: string;
  IsReportsLabel: string;

 //Properties
  TypeLabel: string;
  NodesLabel: string;
  AssocLabel: string;
  NavLabel: string;
  BPMNLabel: string;
  SimLabel: string;

  //ProcTable
  ProcTable: string;
  TaskLabel: string;
  RoleLabel: string;
  RASCILabel: string;
  CommentLabel: string;
  AttachmentLabel: string;
  DOAttachmentLabel: string;
  InputLabel: string;
  OutputLabel: string;
  InputLabelonlyEPC: string;
  OutputLabelonlyEPC: string;

   // Trending, Recent Web part
   NrOfDocumentsToShow: string;
   NoRecentDocuments: string;
   LastUsedMsg: string;
   Loading: string;
   Trending: string;
   Error: string;

   // Relative date strings
   DateJustNow: string;
   DateMinute: string;
   DateMinutesAgo: string;
   DateHour: string;
   DateHoursAgo: string;
   DateDay: string;
   DateDaysAgo: string;
   DateWeeksAgo: string;

// CommandBar
   Search: string;
   Close: string;
   Process: string;
   Navigation: string;
   Hierarchy: string;
   Reports: string;
   Planner: string;
   RoleAssignment: string;
   Bot: string;
   Issues: string;
   Documents: string;
   UsedDocuments: string;
   RecentDocuments: string;
   TrendingDocuments: string;
   People: string;
   Info: string;
   CommandBarEntries: string;
   List: string;

   Diagram: string;
   Label: string;
   Value: string;
   Title: string;
   Home: string;
   Add: string;
   Edit: string;
   Delete: string;
   Submit: string;
   Cancel: string;
   Objects: string;
   Filter: string;
   Column: string;
   Object: string;
   Roles: string;
   Member: string;
   Task: string;
   Completed: string;
   Bucket: string;
   Channel: string;
   ListItemDelete: string;
   PlannerID: string;
   Portal: string;
   BotSecret: string;
   FilterShapeColumn: string;
   FilterPageColumn: string;
   FilterObjectColumn: string;
   FilterModelColumn: string;
   GraphProps: string;

   GoToPage: string;
   GoToObject: string;
   GoToNode: string;
   GoToShape: string;
   LibraryLabel: string;
   LibrarySiteLabel: string;
   SingleDocument: string;
   DirectEdit: string;
   TeamsID: string;
   DefTopic: string;
  }

declare module 'SemTalkStrings' {
  const strings: ISemTalkStrings;
  export = strings;
}
