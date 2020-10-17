import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { ITaxonomyList } from '../components/ITaxonomyList';
import { IListLookupData } from '../components/ILookupData';


export interface IFieldGeneratorLookupData {
    SiteUsers: ISiteUserInfo[];
    Taxonomy: ITaxonomyList;
    ListLookups: IListLookupData;
  }