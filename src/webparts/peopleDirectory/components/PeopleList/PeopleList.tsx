import * as React from 'react';
import { IPeopleListProps } from '.';
import {
  Persona,
  PersonaSize
} from '@fluentui/react/lib/Persona';
import * as strings from 'PeopleDirectoryWebPartStrings';
import styles from './PeopleList.module.scss';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { IPeopleListState } from './IPeopleListState';
import { PeopleCallout } from '../PeopleCallout';
import { IPerson } from '../PeopleDirectory';
//import { PeopleDirectory } from '../PeopleDirectory';
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";
// export interface IPnPPaginationState {
//   allItems: ISPItem[];
//   paginatedItems: ISPItem[];
// }
const pageSize: number = 5;
export class PeopleList extends React.Component<IPeopleListProps, IPeopleListState> {
  [x: string]: any;
  constructor(props: IPeopleListProps) {
    super(props);

    this.state = {
      showCallOut: false,
      calloutElement: null,
      person: null,
      paginatedItems:[],
      allItems:[],

    };

    //this._onPersonaClicked = this._onPersonaClicked.bind(this);
    this._onCalloutDismiss = this._onCalloutDismiss.bind(this);
  }
  public componentDidMount(): void {
    console.log(this.props.people)
    const items: IPerson[] = this.props.people;
    console.log(items);
    
    this.setState({ allItems: this.props.people, paginatedItems: items.slice(0, pageSize) });
    console.log(this.state.allItems)
    console.log(this.state.paginatedItems);
  }
  private _getPage(page: number) {
    // round a number up to the next largest integer.
    const roundupPage = Math.ceil(page);
console.log(page,roundupPage);
    this.setState({
      paginatedItems: this.props.people.slice(roundupPage * pageSize, (roundupPage * pageSize) + pageSize)
    });
    console.log(this.state.paginatedItems)
}

  public render(): React.ReactElement<IPeopleListProps> {
    return (
      <div className='pnPPagination'>
        <div className='container'>
          <div className='row'>
        {
        // this.props.people.length === 0 &&
        //   (this.props.selectedIndex !== 'Search' ||
        //     (this.props.selectedIndex === 'Search' &&
        //       this.props.hasSearchQuery)) &&
              // Show the 'No people found' message if no people have been retrieved
              // and the user either selected a letter in the navigation or issued
              // a search query (but not when navigated to the Search tab without
              // providing a query yet)
          <div className='ms-textAlignCenter'>{strings.NoPeopleFoundLabel}</div>}
          
        {console.log(this.props.people.length) } ;
        {this.props.people.length > 0 &&
          // for each retrieved person, create a persona card with the retrieved
          // information
          //this.props.people.map(p => <Persona primaryText={p.name} secondaryText={p.email} tertiaryText={p.phone} imageUrl={p.photoUrl} imageAlt={p.name} size={PersonaSize.size72} />)
          this.props.paginatedItems.map((p:any,i:number) => {


            const phone: string = p.phone && p.mobile ? `${p.phone}/${p.mobile}`: p.phone ? p.phone: p.mobile;
            // const toggleClassName: string = this.state.toggleClass ? `ms-Icon--ChromeClose ${styles.isClose}` : "ms-Icon--ContactInfo";
            return (
              <div key={i} className={styles.persona_card}>
                <Persona  onClick={this._onPersonaClicked(i, p)} text={p.name} secondaryText={p.email} tertiaryText={phone} imageUrl={p.photoUrl} imageAlt={p.name} size={PersonaSize.size72} />
                <div id={`callout${i}`} onClick={this._onPersonaClicked(i, p)} className={styles.persona}>
                  {/* <p>click here</p> */}
                  <i className="ms-Icon ms-Icon--ContactInfo" aria-hidden="false"></i>
                </div>
              
                
                <p>Hello this is call out</p>
                { this.state.showCallOut && this.state.calloutElement === i && (
                <Callout
                  className={this.state.showCallOut ? styles.calloutShow: styles.callout}
                  gapSpace={8}
                  target={`#callout${i}`}
                  isBeakVisible={true}
                  beakWidth={18}
                  setInitialFocus={true}
                  onDismiss={this._onCalloutDismiss}
                  directionalHint={DirectionalHint.rightCenter}
                  doNotLayer={false}
                >
                  <PeopleCallout person={this.state.person}></PeopleCallout>
                </Callout>
                )}
                
          
              </div>
            );
          })
          
        }

<Pagination
              currentPage={1}
              totalPages={(this.props.people.length / pageSize) - 1}
              // {(this.props.people.length / 5) - 1}
              onChange={(page) => this._getPage(page)}
              limiter={2}
            />

      </div>
     </div>
     </div>
    );
  }

  private _onPersonaClicked = (index: number, person: IPerson) => (_event: any) => {
    this.setState({
      showCallOut: !this.state.showCallOut,
      calloutElement: index,
      person: person
    });
  }

  private _onCalloutDismiss = (_event: any) => {
    this.setState({
      showCallOut: false,
    });
  }
//   private _getPage(page: number){
//     console.log('Page:', page);
//   }
 }
