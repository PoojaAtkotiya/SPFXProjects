import * as React from 'react';
import styles from './DocumentLibView.module.scss';
import { IDocumentLibViewProps, IDocumentLibViewState } from './IDocumentLibViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITheme, getTheme, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { IRectangle } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import Common from '../../common';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';

const theme: ITheme = getTheme();

const ROWS_PER_PAGE = 3;
const MAX_ROW_HEIGHT = 250;
const common: Common = new Common();
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';

export default class DocumentLibView extends React.Component<IDocumentLibViewProps, IDocumentLibViewState> {

  private _columnCount: number;
  private _columnWidth: number;
  private _rowHeight: number;


  constructor(props: IDocumentLibViewProps, state: IDocumentLibViewState) {
    super(props);
    this.state = {
      items: [],
      publishedYears: [],
      catagories: [],
      places: [],
      textTypes: [],

      selectedPublishedYear: '',
      selectedCategory: '',
      selectedPlace: '',
      selectedTextType: '',
      txtTitleValue: '',
      txtAuthorValue: ''
    };

    this.handleAuthorChange = this.handleAuthorChange.bind(this);
    this.handleTitleChanged = this.handleTitleChanged.bind(this);
  }

  public componentDidMount() {
    this.getDocLib();
    this.getDistinctPublishedYear();
    this.getCategoryChoices();
    this.getPlacesChoices();
    this.getTextTypes();
  }

  private getTextTypes() {
    common.getchoices(this.props.siteUrl, "Literature", "TextType").then(resp => {
      var textTypeChoices = [{ key: '', text: "--Select--" }];
      resp.forEach(choice => {
        textTypeChoices.push({ key: choice, text: choice });
      });
      this.setState({
        textTypes: textTypeChoices
      });
    }).catch(error => {
      console.log("error while fatching choices from Office column");
      console.log(error);
    });

  }

  private getPlacesChoices() {
    var listName = "Literature";
    var method = 'get all published year from Literature';
    var query = "?$select=PlaceofPublication"
    common.getDataFromList(this.props.siteUrl, listName, query, method).then(res => {
      if (res.data.value != undefined && res.data.value != null) {
        var dataFiltered = res.data.value;
        var places = [];;
        dataFiltered.forEach(item => {
          places.push(item.PlaceofPublication);
        });
        places = common.removeDuplicatesFromArray(places);
        var placesChoices: IDropdownOption[] = [{ key: '', text: "--Select--" }];
        places.forEach(item => {
          placesChoices.push({ key: item, text: item });
        });
        this.setState({ places: placesChoices });
      }
    }).catch(error => {
      console.log('error while getting data');
      console.log(error);
    });
  }

  private getCategoryChoices() {
    common.getchoices(this.props.siteUrl, "Literature", "Category").then(resp => {
      var categoryChoices = [{ key: '', text: "--Select--" }];
      resp.forEach(choice => {
        categoryChoices.push({ key: choice, text: choice });
      });
      this.setState({
        catagories: categoryChoices
      });
    }).catch(error => {
      console.log("error while fatching choices from Office column");
      console.log(error);
    });
  }

  private getDocLib(): Promise<any> {

    var listName = "Literature";
    var method = 'get items for Literature';
    var query = "?$select=*,FieldValuesAsText&$expand=FieldValuesAsText"
    return common.getDataFromList(this.props.siteUrl, listName, query, method).then(res => {
      if (res.data.value != undefined && res.data.value != null) {
        var dataFiltered = res.data.value;
        this.setState({ items: dataFiltered });
        return this.state.items;
      }
    }).catch(error => {
      console.log('error while getting data');
      console.log(error);
      return null;
    });
  }

  private getDistinctPublishedYear() {
    var listName = "Literature";
    var method = 'get all published year from Literature';
    var query = "?$select=YearPublished"
    common.getDataFromList(this.props.siteUrl, listName, query, method).then(res => {
      if (res.data.value != undefined && res.data.value != null) {
        var dataFiltered = res.data.value;
        var years = [];
        dataFiltered.forEach(item => {
          years.push(item.YearPublished);
        });
        years = years.filter(function (value, index) {
          return years.indexOf(value) === index;
        });
        var yearsChoices: IDropdownOption[] = [{ key: '', text: "--Select--" }];
        years.forEach(item => {
          yearsChoices.push({ key: item, text: item });
        });
        this.setState({ publishedYears: yearsChoices });
      }
    }).catch(error => {
      console.log('error while getting data');
      console.log(error);
    });
  }

  private onSearchClick() {
    var docName = this.state.txtTitleValue;
    var yearPublished = this.state.selectedPublishedYear;
    var category = this.state.selectedCategory;
    var place = this.state.selectedPlace;
    var textType = this.state.selectedTextType;
    var author = this.state.txtAuthorValue;


    this.getDocLib().then(resp => {
      var allItems = this.state.items;
      var filteredItems = allItems;
      if (docName) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.FieldValuesAsText && item.FieldValuesAsText.FileLeafRef && item.FieldValuesAsText.FileLeafRef.toLowerCase().indexOf(docName.toLowerCase()) != -1) {
            return item;
          }
        });
      }
      if (yearPublished) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.YearPublished && item.YearPublished == yearPublished) {
            return item;
          }
        });
      }
      if (category) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.Category && item.Category.length > 0) {
            var catFilteredItems = item.Category.filter(cat => cat == category);
            if (catFilteredItems.length > 0)
              return item;
          }
        });
      }
      if (place) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.PlaceofPublication && item.PlaceofPublication == place) {
            return item;
          }
        });
      }
      if (textType) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.TextType && item.TextType == textType) {
            return item;
          }
        });
      }
      if (author) {
        filteredItems = filteredItems.filter(item => {
          if (item && item.Author0 && item.Author0.toLowerCase().indexOf(author.toLowerCase()) != -1) {
            return item;
          }
        });
      }

      this.setState({
        items: filteredItems
      });
    });

  }

  private onResetClick() {
    this.getDocLib();
    this.setState({
      selectedPublishedYear: "",
      selectedCategory: "",
      selectedPlace: "",
      selectedTextType: "",
      txtTitleValue: "",
      txtAuthorValue: ""
    });
  }

  //#region  on change event of all dropdown and text fields
  private handleYearChanged(event) {
    this.setState({
      selectedPublishedYear: event.key
    });
  }
  private handleCategoryChanged(event) {
    this.setState({
      selectedCategory: event.key
    });
  }
  private handlePlaceChange(event) {
    this.setState({
      selectedPlace: event.key
    });
  }
  private handleTextTypeChanged(event) {
    this.setState({
      selectedTextType: event.key
    });
  }
  private handleTitleChanged(event) {
    this.setState({
      txtTitleValue: event.target.value
    });
  }
  private handleAuthorChange(event) {
    this.setState({
      txtAuthorValue: event.target.value
    });
  }
  //#endregion


  public render(): React.ReactElement<IDocumentLibViewProps> {

    return (
      <div className={styles.row} >
        <div className={styles.column4}>
          <div>
            <TextField id="txtDocTitle" onChange={this.handleTitleChanged} label="Name" value={this.state.txtTitleValue} placeholder="Please Enter Book Name" />
          </div>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Year Published"
              options={this.state.publishedYears}
              id="ddlYearPublished"
              selectedKey={this.state.selectedPublishedYear}
              onChanged={this.handleYearChanged.bind(this)}
              styles={{}}
            />
          </div>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Category"
              options={this.state.catagories}
              onChanged={this.handleCategoryChanged.bind(this)}
              id="ddlCategory"
              styles={{}}
            />
          </div>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Place of Publication"
              options={this.state.places}
              onChanged={this.handlePlaceChange.bind(this)}
              id="ddlPublishedPlace"
              selectedKey={this.state.selectedPlace}
              styles={{}}
            />
          </div>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Text Type"
              options={this.state.textTypes}
              onChanged={this.handleTextTypeChanged.bind(this)}
              id="ddlTextType"
              selectedKey={this.state.selectedTextType}
              styles={{}}
            />
          </div>
          <div>
            <TextField id="txtAuthor" value={this.state.txtAuthorValue} onChange={this.handleAuthorChange} label="Author Name" placeholder="Please Enter Author Name" />
          </div>
          <div style={{padding:'12px' }}>
            <div style={{padding: '4px'}}> <PrimaryButton text="Search" onClick={this.onSearchClick.bind(this)} /></div>
            <div style={{padding: '4px'}}> <DefaultButton text="Reset" onClick={this.onResetClick.bind(this)} /></div>
          </div>
        </div>
        <div className={styles.column8}>

          <FocusZone direction={FocusZoneDirection.vertical}>
            <div className={styles.container} data-is-scrollable={true}>
              <List
                className={styles.listGridExample}
                items={this.state.items}
                getItemCountForPage={this._getItemCountForPage}
                getPageHeight={this._getPageHeight}
                renderedWindowsAhead={4}
                onRenderCell={this._onRenderCell}
              />
            </div>
          </FocusZone>
        </div>
      </div>
    );
  }

  private _getItemCountForPage = (itemIndex: number, surfaceRect: IRectangle): number => {
    if (itemIndex === 0) {
      this._columnCount = 2;//Math.ceil(surfaceRect.width / MAX_ROW_HEIGHT);
      this._columnWidth = Math.floor(surfaceRect.width / this._columnCount);
      this._rowHeight = (this._columnWidth);
    }

    return this._columnCount * ROWS_PER_PAGE;
  };

  private _getPageHeight = (): number => {
    return this._rowHeight * ROWS_PER_PAGE;
  };

  private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {

    return (
      <div
        className={styles.listGridExampleTile}
        data-is-focusable={true}
        role="link"
        style={{
          width: (100 - 6) / this._columnCount + '%'
        }}
      >

        <div className={styles.listGridExampleSizer}>
          <div className={styles.listGridExamplePadder}>
            <div>
              <Link href={item.FieldValuesAsText.FileRef}>
                <div className={styles.listGridExampleLabel}> {item.FieldValuesAsText.FileLeafRef}</div>
              </Link>
              <p>{item.Author0 ? "Author: " + item.Author0 : ""}</p>
              <p>{item.YearPublished ? "Year Published: " + item.YearPublished : ""}</p>
              <p>{item.PlaceofPublication ? " Place Published: " + item.PlaceofPublication : ""}</p>
              <p>{item.TextType ? "Text Type: " + item.TextType : ""}</p>
              <p>{item.Category ? "Category: " + item.Category : ""}</p>
              <p>{item.Description ? common.truncate(item.Description, 110) : ''}</p>
            </div>
          </div >
        </div >
      </div >
    );
  };
}
