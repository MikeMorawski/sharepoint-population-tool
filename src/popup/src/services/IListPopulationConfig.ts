import { IContentTypeInfo, IFieldInfo } from '@pnp/sp/presets/all';
import { IFieldGeneratorLookupData } from './IFieldGeneratorLookupData';

export interface IListPopulationConfig {
    SelectedContentTypes: IContentTypeInfo[];
    FieldGeneratorLookups: IFieldGeneratorLookupData;
    spContext: _spPageContextInfo;
    itemCount: number;
    isExecuting: React.MutableRefObject<boolean | undefined>;
    onProgressUpdate: (count: number) => void;
    onCompletion: () => void;
}