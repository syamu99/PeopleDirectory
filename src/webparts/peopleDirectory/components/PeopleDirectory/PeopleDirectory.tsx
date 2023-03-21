import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './PeopleDirectory.module.scss';
import {
  Spinner,
  SpinnerSize
} from '@fluentui/react/lib/Spinner';
import {
  MessageBar,
  MessageBarType
} from '@fluentui/react/lib/MessageBar';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import {
  IPeopleDirectoryProps,
  IPeopleDirectoryState,
  IPerson,
} from '.';
import { IndexNavigation } from '../IndexNavigation';

import { PeopleList } from '../PeopleList';
import * as strings from 'PeopleDirectoryWebPartStrings';


// const slice: any = require('lodash/slice');

export class PeopleDirectory extends React.Component<IPeopleDirectoryProps, IPeopleDirectoryState> {
  // _onPageUpdate: PageUpdateCallback;
  constructor(props: IPeopleDirectoryProps) {
    super(props);

    this.state = {
      loading: false,
      errorMessage: null,
      selectedIndex: 'A',
      searchQuery: '',
      people: [],
      // totalItems: '',
      // itemsCountPerPage: '0',
      // onPageUpdate: '',
      // currentPage: 0
    };
    
  }
  // const [pagedItems, setPagedItems] = useState<any[]>([]);

  
  
   
  //   const _onPageUpdate = async (pageno?: number) => {
  //     var currentPge = (pageno) ? pageno : currentPage;
  //     var startItem = ((currentPge - 1) * pageSize);
  //     var endItem = currentPge * pageSize;
  //     let filItems = slice( startItem, endItem);
  //     setCurrentPage(currentPge);
  //     setPagedItems(filItems)
  //  };

  private _handleIndexSelect = (index: string): void => {
    // switch the current tab to the tab selected in the navigation
    // and reset the search query
    this.setState({
      selectedIndex: index,
      searchQuery: ''
    },
      function () {
        // load information about people matching the selected tab
        this._loadPeopleInfo(index, null);
      });

  }

  private _handleSearch = (searchQuery: string): void => {
    // activate the Search tab in the navigation and set the
    // specified text as the current search query
    this.setState({
      selectedIndex: 'Search',
      searchQuery: searchQuery
    },
      function () {
        // load information about people matching the specified search query
        this._loadPeopleInfo(null, searchQuery);
      });

  }

  private _handleSearchClear = (): void => {
    // activate the A tab in the navigation and clear the previous search query
    this.setState({
      selectedIndex: 'A',
      searchQuery: ''
    },
      function () {
        // load information about people whose last name begins with A
        this._loadPeopleInfo('A', null);
      });
  }

  /**
   * Loads information about people using SharePoint Search
   * @param index Selected tab in the index navigation or 'Search', if the user is searching
   * @param searchQuery Current search query or empty string if not searching
   */
  private _loadPeopleInfo(index: string, searchQuery: string): void {
    // update the UI notifying the user that the component will now load its data
    // clear any previously set error message and retrieved list of people
    this.setState({
      loading: true,
      errorMessage: null,
      people: []
    });

    const headers: HeadersInit = new Headers();
    // suppress metadata to minimize the amount of data loaded from SharePoint
    headers.append("accept", "application/json;odata.metadata=none");

    // if no search query has been specified, retrieve people whose last name begins with the
    // specified letter. if a search query has been specified, escape any ' (single quotes)
    // by replacing them with two '' (single quotes). Without this, the search query would fail
        // if (query.lastIndexOf('*') !== query.length - 1) {
    //   query += '*';
    // }

    // retrieve information about people using SharePoint People Search
    // sort results ascending by the last name
   // let siteURL= "https://8dk8fn.sharepoint.com/sites/Componentsite/Lists/EmployeeDirectory"

    
//console.log(this.props.webUrl);

// let nexturl= "/_api/search/query?querytext='ListId:bc76b97a-0a40-447f-902c-468e23d3b265AND"+query+"' ";
//+query+ "'&selectproperties='FirstName,LastName,PreferredName,WorkEmail,PictureURL,WorkPhone,MobilePhone,JobTitle,Department,Skills,PastProjects'&sortlist='LastName:ascending'&rowlimit=500";

        // .get(`${siteURL}/_api/search/query?querytext="ListId:4b300120-a80a-41e7-9056-34122fbf6d2b AND givenName: '" + searchQuery + "'"`, SPHttpClient.configurations.v1, {
      // let serachQueryUrl = siteURL+"/_api/web/lists/getbytitle('EmployeeDirectory')/Items?$select=givenName,surname,empFullName,jobTitle,userPrincipleName/EMail&$expand=userPrincipleName &$filter=startswith(givenName,'"+searchQuery+"')` , SPHttpClient.configurations.v1, {
       
      // .get(`${this.props.webUrl}${nexturl}`, SPHttpClient.configurations.v1, {
      
      
      
      
      let query: string;  
      let requiredUrl='';
      if(searchQuery === null)
   {
       requiredUrl="/_api/web/lists/getbytitle('EmpDirectory')/Items?$select=firstName,LastName,PreferredName,WorkEmail,PictureURL,PhoneNumber,MobileNumber,JobTitle,Department,Skills,PastProjects";
   }
   else{
    query = searchQuery === null ? `${index}` : searchQuery.replace(/'/g, `''`);

   requiredUrl="/_api/web/lists/getbytitle('EmpDirectory')/Items?$select=firstName,LastName,PreferredName,WorkEmail,PictureURL,PhoneNumber,MobileNumber,JobTitle,Department,Skills,PastProjects&$filter=startswith(firstName,'${query}')";
   } 
   console.log(searchQuery);
   console.log(query);
console.log(this.props.webUrl+ requiredUrl)
      this.props.spHttpClient 
      .get(`${this.props.webUrl}${requiredUrl}` , SPHttpClient.configurations.v1, {
      headers: headers
      })
      .then((res: SPHttpClientResponse): Promise<void> => {
        console.log(res.json)
        return res.json();
      })
      .then((res: any): void => {
        console.log(res)

        if (res.error) {
          // There was an error loading information about people.
          // Notify the user that loading data is finished and return the
          // error message that occurred
          this.setState({
            loading: false,
            errorMessage: res.error.message
          });
          console.log(res.error.message)
          return;

        }
        console.log(res.value)
      
        if (!res.value) {
          // No results were found. Notify the user that loading data is finished
          this.setState({
            loading: false
          });
          return;
        }
console.log(res);
        // convert the SharePoint People Search results to an array of people
        let people: IPerson[] = res.value?.map((x:any) => {
          return {
            name:x.PreferredName,
            firstName:x.firstName,
            lastName:x.LastName,
            phone:x.PhoneNumber,
            mobile:x.MobileNumber,
            email: x.WorkEmail,
            photoUrl: `${this.props.webUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + x.WorkEmail}`,
            function: x.JobTitle,
            department: x.Department,
            skills: x.Skills,
            projects: x.PastProjects
            // givenName:this.__getValueFromSearchResult('givenName', r.cells)
            // name: this._getValueFromSearchResult('PreferredName', x),
            // firstName: this._getValueFromSearchResult('FirstName', r.Cells),
            // lastName: this._getValueFromSearchResult('LastName', r.Cells),
            // phone: this._getValueFromSearchResult('WorkPhone', r.Cells),
            // mobile: this._getValueFromSearchResult('MobilePhone', r.Cells),
            // email: this._getValueFromSearchResult('WorkEmail', r.Cells),
            // photoUrl: `${this.props.webUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + this._getValueFromSearchResult('WorkEmail', r.Cells)}`,
            // function: this._getValueFromSearchResult('JobTitle', r.Cells),
            // department: this._getValueFromSearchResult('Department', r.Cells),
            // skills: this._getValueFromSearchResult('Skills', r.Cells),
            // projects: this._getValueFromSearchResult('PastProjects', r.Cells)
            // givenName:this.__getValueFromSearchResult('givenName', r.cells)
          };
          
        });
        console.log(people);
        const selectedIndex = this.state.selectedIndex;
console.log(this.state.searchQuery);
        if (this.state.searchQuery === '') {
          // An Index is used to search people.
          //Reduce the people collection if the first letter of the lastName of the person is not equal to the selected index
          people = people.reduce((result: IPerson[], person: IPerson) => {
            if (person. firstName && person. firstName.indexOf(selectedIndex) === 0) {
              result.push(person);
              console.log(people)
               console.log(result)

            }
            return result;
          }, []);
        }

        if (people.length > 0) {
          // notify the user that loading the data is finished and return the loaded information
          this.setState({
            loading: false,
            people: people
          });
        }
        else {
          // People collection could be reduced to zero, so no results
          this.setState({
            loading: false
          });
          return;
        }
      }, (error: any): void => {
        // An error has occurred while loading the data. Notify the user
        // that loading data is finished and return the error message.
        this.setState({
          loading: false,
          errorMessage: error
        });
      })
      .catch((error: any): void => {
        // An exception has occurred while loading the data. Notify the user
        // that loading data is finished and return the exception.
        this.setState({
          loading: false,
          errorMessage: error
        });
      });
  }

  /**
   * Retrieves the value of the particular managed property for the current search result.
   * If the property is not found, returns an empty string.
   * @param key Name of the managed property to retrieve from the search result
   * @param cells The array of cells for the current search result
   */
  // private _getValueFromSearchResult(key: string, cells: ICell[]): string {
  //   for (let i: number = 0; i < cells.length; i++) {
  //     if (cells[i].Key === key) {
  //       return cells[i].Value;
  //     }
  //   }

  //   return '';
  // }

  public componentDidMount(): void {
    // load information about people after the component has been
    // initiated on the page
    
    this._loadPeopleInfo(this.state.selectedIndex, null);
  }

  public render(): React.ReactElement<IPeopleDirectoryProps> {
    const { loading, errorMessage, selectedIndex, searchQuery, people,
     } = this.state;


    return (
      <>
      <div className={styles.peopleDirectory} >
        {!loading &&
          errorMessage &&
          // if the component is not loading data anymore and an error message
          // has been returned, display the error message to the user
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}>{strings.ErrorLabel}: {errorMessage}</MessageBar>
        }
          

        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.onTitleUpdate} />
        <IndexNavigation
          selectedIndex={selectedIndex}
          searchQuery={searchQuery}
          onIndexSelect={this._handleIndexSelect}
          onSearch={this._handleSearch}
          onSearchClear={this._handleSearchClear}
          locale={this.props.locale} />
        {loading &&
          // if the component is loading its data, show the spinner
          <Spinner size={SpinnerSize.large} label={strings.LoadingSpinnerLabel} />
        }
        {!loading &&
          !errorMessage &&
          // if the component is not loading data anymore and no errors have occurred
          // render the list of retrieved people
          <PeopleList
            selectedIndex={selectedIndex}
             hasSearchQuery={searchQuery !== null}
            people={people}
            paginatedItems={people} allItems={people} />
          
        }
        
      </div>
      </>
    );
  }
}


