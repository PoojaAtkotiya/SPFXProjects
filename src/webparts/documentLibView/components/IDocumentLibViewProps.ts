export interface IDocumentLibViewProps {
  description: string;
  siteUrl: string;
  listName: string;
}

export interface IDocumentLibViewState {
  items: any;
  publishedYears: any;
  catagories: any;
  places:any;
  textTypes: any;

  selectedPublishedYear: any;
  selectedCategory:any;
  selectedPlace :any;
  selectedTextType :any;

  txtTitleValue : string;
  txtAuthorValue :string;
}
