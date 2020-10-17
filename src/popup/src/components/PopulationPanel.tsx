import * as React from 'react';
import styles from './PopulationPanel.module.css';
import { sp, IContentTypeInfo, FieldTypes } from "@pnp/sp/presets/all";
import { useEffect } from 'react';
import { Callout, Drawer, AnchorButton, MenuItem, NumericInput, Label, Classes, Dialog, ProgressBar, Intent, NonIdealState } from "@blueprintjs/core";
import { ItemRenderer, MultiSelect } from "@blueprintjs/select";
import "@blueprintjs/icons/lib/css/blueprint-icons.css"
import "@blueprintjs/core/lib/css/blueprint.css"
import axios from 'axios';
import { ITaxonomyList } from './ITaxonomyList';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { IListLookupData } from './ILookupData';
import { ExecuteListPopulation, getEditableFields, isFieldSupported } from '../services/ListPopulation';

// https://github.com/palantir/blueprint/issues/3979 -- strict mode issue

export const PopulationPanel: React.FunctionComponent = () => {
    const [isOpen, setIsOpen] = React.useState(true);
    const handleClose = () => setIsOpen(false);
    const [spLegacy, setSpLegacy] = React.useState<_spPageContextInfo | null>(null);
    const [spListContentTypes, setSpListContentTypes] = React.useState<IContentTypeInfo[]>([]);
    const [spSelectedContentTypes, setSpSelectedContentTypes] = React.useState<IContentTypeInfo[]>([]);
    const [isRunning, setIsRunning] = React.useState(false);
    const [formItemCount, setFormItemCount] = React.useState(100);
    const [progressItemCount, setProgressItemCount] = React.useState(0);
    const isRunningRef = React.useRef<boolean>();
    const [validationErrors, setValidationErrors] = React.useState<string[]>([]);
    const [runComplete, setRunComplete] = React.useState(false);
    const [termsLookup, setTermsLookup] = React.useState<ITaxonomyList>({});
    const [siteUsers, setSiteUsers] = React.useState<ISiteUserInfo[]>([]);
    const [listLookupIds, setListLookupIds] = React.useState<IListLookupData>({});



    // Loads metadata of the content types to prepare for population
    useEffect(() => {
        const loadData = async (ctx: _spPageContextInfo) => {
            if (ctx != null) {
                sp.setup({ sp: { headers: { Accept: "application/json;odata=verbose" }, baseUrl: ctx.webAbsoluteUrl } });

                const taxonomyLookupDictionary: ITaxonomyList = {} as ITaxonomyList;
                const lookupFieldDictionary: IListLookupData = {} as IListLookupData;

                // Load content types and fields
                var cts = await sp.web.getList(ctx.listUrl).contentTypes.expand("Fields").get();
                setSpListContentTypes(cts);

                // Field preprocessing
                await cts.forEach(async ct => {
                    await getEditableFields(ct).forEach(async field => {
                        // Taxonomy preprocessing
                        if (["TaxonomyFieldTypeMulti", "TaxonomyFieldType"].includes(field.TypeAsString)) {
                            const termId: string = (field as any).TermSetId;
                            const termGroupId = (await axios.get(ctx.siteServerRelativeUrl + "/_api/v2.0/termStore/sets/" + termId)).data.groupId;
                            // Improvement: Get tree rather than first tier children
                            let allTax = await sp.termStore.termGroups.getById(termGroupId).termSets.getById(termId).children();
                            taxonomyLookupDictionary[termId] = allTax;
                        }
                        // Lookup preprocessing
                        else if (field.FieldTypeKind === FieldTypes.Lookup) {
                            var lookupItems = await sp.web.lists.getById((field as any).LookupList).items.getAll();
                            lookupFieldDictionary[field.InternalName] = lookupItems.map(x => x.Id)
                        }
                    });
                });
                setTermsLookup(taxonomyLookupDictionary);
                setListLookupIds(lookupFieldDictionary);

                // Preload user choices for people or group fields
                let siteUsers = await sp.web.siteUsers.get();
                // Improvement: Load groups as well for more diverse data
                siteUsers = siteUsers.filter(x => { return x.PrincipalType === 1 && !!x.UserPrincipalName && x.Id });
                setSiteUsers(siteUsers);
            }
        }

        chrome.runtime.onMessage.addListener(function (message) {
            setSpLegacy(message.spLegacyCtx);
            loadData(message.spLegacyCtx);
        });
    });

    const onPopulateClicked = async () => {
        if (validateForm()) {
            isRunningRef.current = true;
            setRunComplete(false);
            setIsOpen(false);
            setIsRunning(true);
        }
    }

    const validateForm = (): boolean => {
        const errors = [];
        if (spSelectedContentTypes.length === 0) {
            errors.push("At least one content type must be selected");
        }
        setValidationErrors(errors);
        return errors.length === 0;
    }

    useEffect(() => {
        var populate = async () => {
            if (!isRunning || spLegacy === null)
                return;

            ExecuteListPopulation({
                FieldGeneratorLookups: {
                    ListLookups: listLookupIds,
                    SiteUsers: siteUsers,
                    Taxonomy: termsLookup
                },
                spContext: spLegacy,
                SelectedContentTypes: spSelectedContentTypes,
                itemCount: formItemCount,
                isExecuting: isRunningRef,
                onCompletion: () => { setRunComplete(true); },
                onProgressUpdate: (iter: number) => { setProgressItemCount(iter); }
            });
        }

        populate();
    }, [isRunning, spSelectedContentTypes, spLegacy, formItemCount, listLookupIds, siteUsers, termsLookup]);

    const renderDropDownCtItem: ItemRenderer<IContentTypeInfo> = (ctInfo, { modifiers, handleClick }) => {
        if (!modifiers.matchesPredicate) {
            return null;
        }
        return (
            <MenuItem
                active={modifiers.active}
                key={ctInfo.Id.StringValue}
                label={ctInfo.Group}
                onClick={handleClick}
                text={`${ctInfo.Name}`}
                shouldDismissPopover={false}
            />
        );
    };


    const onContentTypeSelection = (item: IContentTypeInfo, event?: React.SyntheticEvent<HTMLElement, Event> | undefined) => {
        // Add to selected content types if not already selected
        const isCtAlreadySelected = spSelectedContentTypes.indexOf(item) !== -1;
        if (!isCtAlreadySelected) {
            setSpSelectedContentTypes([...spSelectedContentTypes, item]);
        }
    };

    var ctSelectedRenderer = (ct: IContentTypeInfo) => {
        return ct.Name;
    };

    var stopPopulation = () => {
        isRunningRef.current = false;
        setIsRunning(false);
    }

    return (
        <React.Fragment>
            <Drawer
                icon="add-to-artifact"
                onClose={handleClose}
                title="List Population Tool"
                isOpen={isOpen}
                className={styles.PopulationPanel}
                size={Drawer.SIZE_SMALL}
            >
                <div className={styles.contentWrapper}>
                    {(spLegacy != null && spLegacy.listUrl != null && spLegacy.listUrl.length > 0) ?
                        <div className={styles.panelMainContent}>

                            {(spLegacy.listBaseTemplate !== 119 && spLegacy.listBaseTemplate !== 100) &&

                                <Callout className={styles.contextCalloutWarning}
                                    title="Limited Support" intent={Intent.DANGER}
                                    icon={"search-template"}>
                                    Currently this utility only supports <strong>Site Pages</strong> and <strong>Basic Lists</strong>. Functionality at this point may not work as intended.
                                </Callout>
                            }

                            Configure the content types of items to generate and click populate list to add randomized data

                            <h3 className={styles.sectionHeading}>Content Types</h3>
                            <span>Select the list content types to generate as part of this execution plan.</span>
                            <MultiSelect
                                itemRenderer={renderDropDownCtItem}
                                tagRenderer={ctSelectedRenderer}
                                onItemSelect={onContentTypeSelection}
                                items={spListContentTypes}
                                selectedItems={spSelectedContentTypes}
                                placeholder="Select content types..."
                                tagInputProps={{
                                    onRemove: (_tag: string, index: number) => {
                                        var currentSelection = [...spSelectedContentTypes];
                                        currentSelection.splice(index, 1);
                                        setSpSelectedContentTypes(currentSelection);
                                    },
                                }}
                                popoverProps={{ targetTagName: 'div' }}
                                className={styles.multi}
                            >
                            </MultiSelect>

                            <h3 className={styles.sectionHeading}>Field Overview</h3>
                            <div>Fields for each content type which are writable, non-hidden, and visible in edit form:</div>


                            {spSelectedContentTypes.length === 0 &&
                                <p className={styles.noFieldsMessage}>Select a content type above to review fields.</p>
                            }
                            {spSelectedContentTypes.map((ct, i) => {
                                return (
                                    <div className={styles.fieldInfoItem} key={ct.Id.StringValue}>
                                        <h4 className={styles.ctNameLabel}>{ct.Name}</h4>
                                        <ul className={styles.fieldList}>
                                            {getEditableFields(ct).map((field, i) => {
                                                return (<li key={field.Id}>
                                                    {field.Title}
                                                    {(field.StaticName !== "FileLeafRef" && !isFieldSupported(field)) && //FileLeafRef name indirectly supported
                                                        <span className={styles.noFieldSupport}>[Unsupported]</span>
                                                    }
                                                </li>)
                                            })}
                                        </ul>
                                    </div>
                                )
                            })}

                            <h3 className={styles.sectionHeading}>Provisioning Settings</h3>

                            <Label>
                                Item Count
                                <NumericInput
                                    value={formItemCount}
                                    onValueChange={(valueAsNumber: number, _valueAsString: string) => { setFormItemCount(valueAsNumber); }}
                                    stepSize={100}
                                    majorStepSize={1000}
                                    min={0} />
                            </Label>

                            <Callout className={styles.contextCalloutWarning} title="Sharepoint Context" intent={Intent.WARNING} hidden={spLegacy != null && decodeURI(window.location.pathname).toLowerCase().indexOf(spLegacy?.listUrl.toLowerCase()) > -1}>
                                The current SP Context may not align with the current list shown.
                                Please refresh to ensure that the correct list will be targeted.
                                <br />
                                <br />
                                <strong>Target:</strong> {spLegacy?.listUrl}
                            </Callout>

                            <AnchorButton text="Populate List" className="bp3-intent-primary" onClick={onPopulateClicked} />

                            <ul hidden={validationErrors.length === 0}>
                                {validationErrors.map((error) => {
                                    return (
                                        <li>{error}</li>
                                    )
                                })}
                            </ul>
                        </div>
                        :
                        <div>
                            <NonIdealState
                                icon={"search-template"}
                                title="List Context Missing"
                                description="No list found on this page. Ensure you are viewing the contents of a list, if you are, try refreshing."
                            />
                        </div>
                    }
                    <div className={styles.footer}>
                        Developed by <a href="http://www.migee.com" target="_blank" rel="noopener noreferrer">Mike Morawski</a><br />
                        <a href="https://github.com/MikeMorawski/sharepoint-population-tool" target="_blank" rel="noopener noreferrer">View on GitHub</a>
                    </div>
                </div>
            </Drawer>

            <Dialog
                icon="info-sign"
                title="Populating List"
                isOpen={isRunning}
                hasBackdrop={false}
                onClose={stopPopulation}
                canOutsideClickClose={false}
                canEscapeKeyClose={false}
            >
                <div className={Classes.DIALOG_BODY}>
                    <p>
                        List population is in progress. Do not navigate away from this page or close this window.
                        To review items as they are being generated click <a href={window.location.href} target="_blank" rel="noopener noreferrer">this link</a> to open this list in a new tab.
                    </p>
                    <ProgressBar intent={runComplete ? Intent.SUCCESS : Intent.PRIMARY} value={runComplete ? 1 : progressItemCount / formItemCount} ></ProgressBar>
                    <div>Processing item {progressItemCount} of {formItemCount}</div>
                </div>
                <div className={Classes.DIALOG_FOOTER}>
                    <div className={Classes.DIALOG_FOOTER_ACTIONS}>
                        {!runComplete ?
                            <AnchorButton onClick={stopPopulation} intent={Intent.PRIMARY}>
                                <span>Stop Population</span>
                            </AnchorButton>
                            :
                            <AnchorButton onClick={stopPopulation} intent={Intent.SUCCESS}>
                                <span>Close</span>
                            </AnchorButton>
                        }
                    </div>
                </div>
            </Dialog>
        </React.Fragment >
    );
};