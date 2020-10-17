import { IFieldInfo, FieldTypes } from "@pnp/sp/presets/all";
import { addDays } from "date-fns";
import { IFieldGeneratorLookupData } from "./IFieldGeneratorLookupData";
import * as _ from "lodash";
const randomWords = require('random-words');

// Generates random sample data for each individual field type.
export const randomFieldData = (field: IFieldInfo, lookups: IFieldGeneratorLookupData) => {
    let fieldAny = field as any;
    let maxVal: number;
    let minVal: number;

    switch (field.FieldTypeKind) {
        case FieldTypes.Note:
            return randomWords({ min: 1, max: 500, join: ' ' });
        case FieldTypes.Text:
            let maxLength = fieldAny.MaxLength;
            return randomWords({ min: 1, max: maxLength / 6, join: ' ' }).substring(0, maxLength);
        case FieldTypes.Boolean:
            return Math.round(Math.random()) === 1;
        case FieldTypes.Choice:
            return _.sample(fieldAny.Choices.results);
        case FieldTypes.Currency:
            maxVal = fieldAny.MaximumValue;
            minVal = fieldAny.MinimumValue;
            return randomIntFromInterval(minVal, maxVal) / 100;
        case FieldTypes.DateTime:
            return addDays(new Date(), Math.random() * 365);
        case FieldTypes.Integer:
            maxVal = fieldAny.MaximumValue;
            minVal = fieldAny.MinimumValue;
            return randomIntFromInterval(minVal, maxVal);
        case FieldTypes.MultiChoice:
            if (fieldAny.Choices.results.length > 0) {
                const selectedChoices = _.sampleSize(fieldAny.Choices.results, Math.floor(Math.random() * fieldAny.Choices.results.length) + 1);
                return {
                    results: selectedChoices
                }
            } else {
                return null;
            }
        case FieldTypes.Number:
            maxVal = fieldAny.MaximumValue;
            minVal = fieldAny.MinimumValue;
            return randomIntFromInterval(minVal, maxVal);
        case FieldTypes.URL:
            return {
                Description: randomWords({ min: 1, max: 10, join: ' ' }),
                Url: "https://www." + randomWords({ min: 1, max: 2, join: '' }) + ".com"
            }
        case FieldTypes.User:
            let userLookupsExist = lookups.SiteUsers !== undefined && lookups.SiteUsers.length > 0;
            if (field.TypeAsString === "UserMulti") {
                const user = userLookupsExist ? _.sampleSize(lookups.SiteUsers, Math.floor(Math.random() * lookups.SiteUsers.length + 1)).map(x => x.Id) : null;
                if (user !== null && user.length > 0) {
                    return { 'results': user };
                } else {
                    return null;
                }
            } else {
                const user = userLookupsExist ? _.sample(lookups.SiteUsers)!.Id : null;
                return user;
            }
        case FieldTypes.Lookup:
            if (field.TypeAsString === "LookupMulti") {
                const curLookup = lookups.ListLookups[field.InternalName];
                const randomLookups = (curLookup !== undefined && curLookup.length > 0) ? _.sampleSize(curLookup, Math.floor(Math.random() * curLookup.length + 1)) : null;
                if (randomLookups !== null && randomLookups.length > 0) {
                    return { 'results': randomLookups };
                } else {
                    return null;
                }
            } else {
                return _.sample(lookups.ListLookups[field.InternalName]);
            }
        case FieldTypes.Invalid:
            if (field.TypeAsString === "TaxonomyFieldTypeMulti") {
                const terms = _.sampleSize(lookups.Taxonomy[fieldAny.TermSetId], Math.floor(Math.random() * lookups.Taxonomy[fieldAny.TermSetId].length) + 1);
                let termsString = "";
                terms.forEach(x => {
                    termsString += `-1;#${x.labels.find(x => x.isDefault)!.name}|${x.id};#`;
                });
                return termsString;
            } else if (field.TypeAsString === "TaxonomyFieldType") {
                const term = _.sample(lookups.Taxonomy[fieldAny.TermSetId]);
                if (term !== undefined) {
                    return {
                        "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                        "Label": term.labels.find(x => x.isDefault)!.name,
                        'TermGuid': term.id,
                        'WssId': '-1'
                    };
                }
            }
            return null;
        default:
            return null;
    }
}

// Generates an integer between/on min and max values /w additional failsafe logic.
var randomIntFromInterval = (min: number, max: number) => { // min and max included 
    let randomInt = Math.floor(Math.random() * (max - min + 1) + min);

    // Improvement: Replace min/max with the field itself for better control.

    //Failsafe if for any reason we get infinity
    if (!isFinite(randomInt)) {
        randomInt = Math.floor(Math.random() * (10000 - (-10000) + 1) + (-10000));
    }

    return randomInt;
}