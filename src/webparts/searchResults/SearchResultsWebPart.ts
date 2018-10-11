import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Text, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  PropertyPaneSlider,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupOption,
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneCustomField,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import {Environment, EnvironmentType } from "@microsoft/sp-core-library";
import * as strings from 'SearchsWebPartStrings';
import SearchContainer from './components/SearchResultsContainer/SearchResultsContainer';
import ISearchContainerProps from './components/SearchResultsContainer/ISearchResultsContainerProps';
import { ISearchResultsWebPartProps } from './ISearchResultsWebPartProps';
import ISearchService from '../../services/SearchService/ISearchService';
import MockSearchService from '../../services/SearchService/MockSearchService';
import SearchService from '../../services/SearchService/SearchService';
import ITaxonomyService from '../../services/TaxonomyService/ITaxonomyService';
import MockTaxonomyService from '../../services/TaxonomyService/MockTaxonomyService';
import TaxonomyService from '../../services/TaxonomyService/TaxonomyService';
import { Placeholder, IPlaceholderProps } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import LocalizationHelper from '../../helpers/LocalizationHelper';
import ResultsLayoutOption from '../../models/ResultsLayoutOption';
import TemplateService from '../../services/TemplateService/TemplateService';
import { update, isEmpty } from '@microsoft/sp-lodash-subset';
import MockTemplateService from '../../services/TemplateService/MockTemplateService';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import { IDynamicDataSource } from '@microsoft/sp-dynamic-data';

declare var System: any;

const LOG_SOURCE: string = '[SearchResultsWebPart_{0}]';

export default class SearchResultsWebPart extends BaseClientSideWebPart<ISearchResultsWebPartProps> {
  private _searchService: ISearchService;
  private  _taxonomyService: ITaxonomyService;
  private _templateService: BaseTemplateService;
  private _useResultService: boolean;
  private _queryKeywords: string;
  private _source: IDynamicDataSource;
  private _domElement: HTMLElement;
  private _propertyPage: null;

  private _lastSourceId: string = undefined;
  private _lastPropertyId: string = undefined;

  // Template to display at render time
  private _templateContentToDisplay: string;

  constructor() {
    super();
    this._parseRefiners = this._parseRefiners.bind(this);
  }

  /**
   * Resolves the connected data sources
   * Useful in the case when the data source comes from an extension,
   * the id is regenerated every time the page is refreshed causing the property pane configuration be lost
   */
  private _initDynamicDataSource() {
    if (this.properties.dynamicDataSourceId
      && this.properties.dynamicDataSourcePropertyId
      && this.properties.dynamicDataSourceComponentId) {
      this.source = this.context.dynamicDataProvider.tryGetSource(this.properties.dynamicDataSourceId);
      let sourceId = undefined;

      if (this._source) {
        sourceId = this._source.id;
      } else {
        this._source = this._tryGetSourceByComponentId(this.properties.dynamicDataSourceComponentId);
        sourceId = this._source ? this._source.id : undefined;
      }

      if (sourceId) {
        this.context.dynamicDataProvider.registerPropertyChanged(sourceId, this.properties.dynamicDataSourcePropertyId, this.render);

        this.properties.dynamicDataSourceId = sourceId;
        this._lastSourceId = this.properties.dynamicDataSourceId;
        this._lastPropertyId = this.properties.dynamicDataSourcePropertyId;

        if (this.renderedOnce) {
          this.render();
        }
      }
    }
  }

  private _tryGetSourceByComponentId(dataSourceComponentId: string): IDynamicDataSource {
    const resolvedDataSource = this.context.dynamicDataProvider.getAvailableSources()
      .filter((item) => {
        if (item.metadata.componentId) {
          if (item.metadata.componentId.localeCompare(dataSourceComponentId) === 0) {
            return item;
          }
        }
      });

    if(resolvedDataSource.length > 0) {
      return resolvedDataSource[0];
    } else {
      Log.verbose(Text.format(LOG_SOURCE, "_tryGetSourceByComponentId()"), `Unable to find dynamic data source with componentId '${dataSourceComponentId}'`);
      return undefined;
    }
  }

  /**
   * Determines the group fields for the search settings options inside the property pane
   */
  private _getSearchSettingsFields(): IPropertyPaneField<any>[] {
    // Sets up search settings fields
    const searchSettingsFields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('queryTemplate', {
        label: strings.QueryTemplateFieldLabel,
        value: this.properties.queryTemplate,
        multiline: true,
        resizable: true,
        placeholder: strings.SearchQueryPlaceHolderText,
        deferredValidationTime: 300,
        disabled: this._useResultSource,
      }),
      PropertyPaneTextField('resultSourceId', {
        label: strings.ResultSourceIdLabel,
        multiline: false,
        onGetErrorMessage: this.validateSourceId.bind(this),
        deferredValidationTime: 300
      }),
      PropertyPaneTextField('sortList', {
        label: strings.SortList,
        description: strings.SortListDescription,
        multiline: false,
        resizable: true,
        value: this.properties.sortList,
        deferredValidationTime: 300
      }),
      PropertyPaneToggle('enableQueryRules', {
        label: strings.EnableQueryRulesLabel,
        checked: this.properties.enableQueryRules,
      }),
      PropertyPaneTextField('selectedProperties', {
        label: strings.SelectedPropertiesFieldLabel,
        description: strings.SelectedPropertiesFieldDescription,
        multiline: true,
        resizable: true,
        value: this.properties.selectedProperties,
        deferredValidationTime: 300
      }),
      PropertyPaneTextField('refiners', {
        label: strings.RefinersFieldLabel,
        description: strings.RefinersFieldDescription,
        multiline: true,
        resizable: true,
        value: this.properties.refiners,
        deferredValidationTime: 300,
      })
    ];

    return searchSettingsFields;
  }

  /**
   * Determines the group fields for the search query options inside the property pane
   */
  private _getSearchQueryFields(): IPropertyPaneField<any>[] {
    // Sets up search query fields
    let searchQueryConfigFields: IPropertyPaneField<any>[] = [
      PropertyPaneCheckbox('useSearchBoxQuery', {
        checked: false,
        text: strings.UseSearchBoxQueryLabel,
      })
    ];

    if (this.properties.useSearchBoxQuery) {
      const sourceOptions: IPropertyPaneDropdownOption[] =
        this.context.dynamicDataProvider.getAvailableSources().map(source => {
          return {
            key: source.id,
            text: source.metadata.title
          };
        }).filter(item => item.key.localeCompare("PageContext") !== 0);

      const selectedSource: string = this.properties.dynamicDataSourceId;

      let propertyOptions: IPropertyPaneDropdownOption[] = [];
      if (selectedSource) {
        const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(selectedSource);
        if (source) {
          propertyOptions = source.getPropertyDefinitions().map(prop => {
            return {
              key: prop.id,
              text: prop.title
            };
          });
        }
      }

      searchQueryConfigFields = searchQueryConfigFields.concat([
        PropertyPaneDropdown('dynamicDataSourceId', {
          label: "Source",
          options: sourceOptions,
          selectedKey: this.properties.dynamicDataSourceId,
        }),
        PropertyPaneDropdown('dynamicDataSourcePropertyId', {
          disabled: !this.properties.dynamicDataSourceId,
          label: "Source property",
          options: propertyOptions,
          selectedKey: this.properties.dynamicDataSourcePropertyId
        }),
      ]);
    } else {
      searchQueryConfigFields.push(
        PropertyPaneTextField('queryKeywords', {
          label: strings.SearchQueryKeywordsFieldLabel,
          description: strings.SearchQueryKeywordsFieldDescription,
          value: this.properties.useSearchBoxQuery ? '' : this.properties.queryKeywords,
          multiline: true,
          resizable: true,
          placeholder: strings.SearchQueryPlaceHolderText,
          onGetErrorMessage: this._validateEmptyField.bind(this),
          deferredValidationTime: 500,
          disabled: this.properties.useSearchBoxQuery
        })
      );
    }

    searchQueryConfigFields.push(
      PropertyPaneLabel('', { text: '' }),
      PropertyPaneSlider('maxResultsCount', {
        label: strings.MaxResultsCount,
        max: 50,
        min: 1,
        showValue: true,
        step: 1,
        value: 50,
      })
    );

    return searchQueryConfigFields;
  }

  /**
   * Determines the group fields for styling options inside the property pane
   */
  private _getStylingFields(): IPropertyPaneField<any>[] {
    const layoutOptions = [
      {
        iconProps: {
          officeFabricIconFontName: 'List'
        },
        text: strings.ListLayoutOption,
        key: ResultsLayoutOption.List,
      },
      {
        iconProps: {
          officeFabricIconFontName: 'Tiles'
        },
        text: strings.TilesLayoutOption,
        key: ResultsLayoutOption.Tiles
      },
      {
        iconProps: {
          officeFabricIconFontName: 'Code'
        },
        text: strings.CustomLayoutOption,
        key: ResultsLayoutOption.Custom
      }
    ] as IPropertyPaneChoiceGroupOption[];

    const canEditTemplate = this.properties.externalTemplateUrl && this.properties.selectedLayout == ResultsLayoutOption.Custom ? false : true;

    const pp: IPropertyPaneCustomFieldProps = {
      onRender: (elem: HTMLElement): void => {
        elem.innerHTML = `<div class="ms-font-xs ms-fontColor-neutralSecondary">${strings.HandlebarsHelpersDescription}</div>`;
      },
      key: "HandelbarsDescription"
    };

    // Sets up styling fields
    let stylingFields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('webPartTitle', {
        label: strings.WebPartTitle
      }),
      PropertyPaneToggle('showBlank', {
        label: strings.ShowBlankLabel,
        checked: this.properties.showBlank,
      }),
      PropertyPaneToggle('showResultsCount', {
        label: strings.ShowResultsCountLabel,
        checked: this.properties.showResultsCount,
      }),
      PropertyPaneToggle('showPaging', {
        label: strings.ShowPagingLabel,
        checked: this.properties.showPaging,
      }),
      PropertyPaneChoiceGroup('selectedLayout', {
        label: 'Results layout',
        options: layoutOptions
      }),
      new this._propertyPage.PropertyPaneTextDialog('inlineTemplateText', {
        dialogTextFieldValue: this._templateContentToDisplay,
        onPropertyChange: this._onCustomPropertyPaneChange.bind(this),
        disabled: !canEditTemplate,
        strings: {
          cancelButtonText: strings.CancelButtonText,
          dialogButtonLabel: strings.DialogButtonLabel,
          dialogButtonText: strings.DialogButtonText,
          dialogTitle: strings.DialogTitle,
          saveButtonText: strings.SaveButtonText
        }
      }),
      PropertyPaneToggle('useHandlebarsHelpers', {
        label: "Handlebars Helpers",
        checked: this.properties.useHandlebarsHelpers
      }),
      PropertyPaneCustomField(pp)
    ];

    // Only show the template external URL for 'Custom' option
    if (this.properties.selectedLayout === ResultsLayoutOption.Custom) {
      stylingFields.push(PropertyPaneTextField('externalTemplateUrl', {
        label: strings.TemplateUrlFieldLabel,
        placeholder: strings.TemplateUrlPlaceholder,
        deferredValidationTime: 500,
        onGetErrorMessage: this._onTemplateUrlChange.bind(this)
      }));
    }

    return stylingFields;
  }

  /**
   * Opens the Web Part property pane
   */
  private _setupWebPart() {
    this.context.propertyPane.open();
  }

  /**
   * Checks if a field if empty or not
   * @param value the value to check
   */
  private _validateEmptyField(value: string): string {
    if (!value) {
      return strings.EmptyFieldErrorMessage;
    }
    return '';
  }

  /**
   * Ensures the result source id value is a valid GUID
   * @param value the result source id
   */
  private validateSourceId(value: string): string {
    if (value.length > 0) {
      if (!/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/.test(value)) {
        this._useResultSource = false;
        return strings.InvalidResultSourceIdMessage;
      } else {
        this._useResultSource = true;
      }
    } else {
      this._useResultSource = false;
    }

    return '';
  }

  /**
   * Parses refiners from the property pane value by extracting the refiner managed property and its label in the filter panel.
   * @param rawValue the raw value of the refiner
   */
  private _parseRefiners(rawValue: string): { [key: string]: string } {
    let refiners = {};

    // Get each configuration
    let refinerKeyValuePair = rawValue.split(',');

    if (refinerKeyValuePair.length > 0) {
      refinerKeyValuePair.map((e) => {
        const refinerValues = e.split(':');
        switch (refinerValues.length) {
          case 1:
            refiners[refinerValues[0]] = refinerValues[0];
            break;
          case 2:
            refiners[refinerValues[0]] = refinerValues[1].replace(/^'(.*)'$/, '$1');
            break;
        }
      });
    }
    return refiners;
  }

  /**
   * Get the correct results template content according to the property pane current configuration
   * @returns the template content as a string
   */
  private async _getTemplateContent(): Promise<void> {
    let templateContent = null;
    switch (this.properties.selectedLayout) {
      case ResultsLayoutOption.List:
        templateContent = TemplateService.getListDefaultTemplate();
        break;
      case ResultsLayoutOption.Tiles:
        templateContent = TemplateService.getTilesDefaultTemplate();
        break;
      case ResultsLayoutOption.Custom:
        if (this.properties.externalTemplateUrl) {
          templateContent = await this._templateService.getFileContent(this.properties.externalTemplateUrl);
        } else {
          templateContent = this.properties.inlineTemplateText ? this.properties.inlineTemplateText : TemplateService.getBlankDefaultTemplate();
        }
        break;
      default:
        break;
    }

    this._templateContentToDisplay = templateContent;
  }
  public render(): void {
    const element: React.ReactElement<ISearchResultsProps > = React.createElement(
      SearchResults,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
