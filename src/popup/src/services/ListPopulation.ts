import { IListPopulationConfig } from "./IListPopulationConfig";
import { IContentTypeInfo, FieldTypes, sp, IFieldInfo } from "@pnp/sp/presets/all";
import { randomFieldData } from "./FieldGenerator";
import * as _ from "lodash";

const randomWords = require('random-words');

const supportedFieldTypes = [
    FieldTypes.Note,
    FieldTypes.Text,
    FieldTypes.Boolean,
    FieldTypes.Currency,
    FieldTypes.DateTime,
    FieldTypes.Integer,
    FieldTypes.Choice,
    FieldTypes.MultiChoice,
    FieldTypes.Number,
    FieldTypes.URL,
    FieldTypes.User,
    FieldTypes.Lookup
];
const supportedTypeAsStrings = [
    "TaxonomyFieldTypeMulti",
    "TaxonomyFieldType"
];

export const ExecuteListPopulation = async (cfg: IListPopulationConfig) => {
    //Get content types to understand what data to create
    for (let i = 0; i < cfg.itemCount; i++) {
        if (!cfg.isExecuting.current) {
            console.log("Population cancelled by user");
            break;
        }

        cfg.onProgressUpdate(i + 1);

        //Randomly select CT
        const randomCT = _.sample(cfg.SelectedContentTypes) as IContentTypeInfo;

        let itemHash = { ContentTypeId: randomCT.StringId } as any;
        let isFolder = randomCT.StringId.toLowerCase().startsWith("0x0120");

        //Populate new item
        let fields = getEditableFields(randomCT);

        //fields = fields.slice(0, 10);
        //console.log(fields);

        fields.forEach(f => {
            let data = randomFieldData(f, cfg.FieldGeneratorLookups);

            if (data != null) {
                let fieldDataKey = f.InternalName;
                //Handle multi taxonomy fields which require targeting another field
                if (f.TypeAsString === "TaxonomyFieldTypeMulti") {
                    //Find corresponding field internal name
                    const correspondingTaxonomyStaticName = getInternalTaxonomyField(randomCT, f.Title);
                    if (correspondingTaxonomyStaticName) {
                        fieldDataKey = correspondingTaxonomyStaticName;
                    }
                }
                //Handle User Field and Lookup updates by Appending ID to field name
                if (f.FieldTypeKind === FieldTypes.User || f.FieldTypeKind === FieldTypes.Lookup) {
                    fieldDataKey = fieldDataKey + "Id";
                }


                itemHash[fieldDataKey] = data;
            }
        });

        if (isFolder) {
            await sp.web.getList(cfg.spContext.listUrl).rootFolder.folders.add(randomWords({ exactly: 2, maxLength: 10, join: ' ' }))
        }
        else {
            const lTemplate = cfg.spContext.listBaseTemplate;
            if (lTemplate === 119) { //Site Pages
                var page = await (await sp.web.addClientsidePage(randomWords({ exactly: 2, maxLength: 10, join: ' ' }), undefined, "Article"));
                var pageId = (page as any).json.Id
                await sp.web.getList(cfg.spContext.webServerRelativeUrl + "/sitepages").items.getById(pageId).update(itemHash);
                page.save(true);
            } else if (lTemplate === 101) { //Doc lib

            } else if (lTemplate === 100) { //Generic List
                await sp.web.getList(cfg.spContext.listUrl).items.add(itemHash);
            }
        }
    }

    cfg.onCompletion();
}

export const getEditableFields = (ct: IContentTypeInfo) => {
    return ((ct as any).Fields.results as IFieldInfo[]).filter(
        x => !x.Hidden
            && x.TypeAsString !== "Computed"
            && x.TypeAsString !== "Calculated"
            && !x.ReadOnlyField
            && x.SchemaXml.toLowerCase().indexOf('showineditform="false"') === -1
            && (x.Group !== "_Hidden" || ["Title","Description"].includes(x.InternalName))
            && !["Modified_x0020_By", "FileLeafRef", "Created_x0020_By"].includes(x.InternalName)
    );
}

export const isFieldSupported = (f: IFieldInfo) => {
    return supportedFieldTypes.includes(f.FieldTypeKind) || supportedTypeAsStrings.includes(f.TypeAsString);
}

const getInternalTaxonomyField = (ct: IContentTypeInfo, primaryTaxonomyFieldName: string) => {
    return ((ct as any).Fields.results as IFieldInfo[]).find(x => x.Title === primaryTaxonomyFieldName + "_0")?.StaticName;
}