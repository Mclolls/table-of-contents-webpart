import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField
  //,   PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';
import { IPropertyPaneField } from '@microsoft/sp-property-pane';
import { PropertyFieldSwatchColorPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSwatchColorPicker';
//import { PropertyFieldSwatchColorPicker, PropertyFieldSwatchColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldSwatchColorPicker';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import { TableOfContents } from './components/TableOfContents';

export interface ITableOfContentsWebPartProps {
  showH2: boolean;
  showH3: boolean;
  showH4: boolean;
  listType: 'bulleted' | 'numbered';
  listStyle: 'disc' | 'circle' | 'square' | 'none' | 'decimal' | 'lower-alpha' | 'upper-alpha' | 'lower-roman' | 'upper-roman';
  title?: string;
  titleColorMode: 'theme' | 'standard' | 'custom'; // New
  itemsColorMode: 'theme' | 'standard' | 'custom'; // New
  titleColor?: string;
  itemsColor?: string;
  columns: 1 | 2 | 3;
   minColumnWidth: 150 | 240 | 280 | 300 | 320;
  debugHeadings?: boolean;
  scope: 'page' | 'section';
  groupBySection?: boolean;
  showSectionHeadings?: boolean;
    // NEW: allow the editor to choose the title tag
  titleLevel?: 'h2' | 'h3' | 'h4';
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {
  private _themeProvider!: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public render(): void {
    const levels = [
      this.properties.showH2 ? 2 : null,
      this.properties.showH3 ? 3 : null,
      this.properties.showH4 ? 4 : null
    ].filter((l): l is number => l !== null);

    const element = React.createElement(TableOfContents, {
      levels,
      listType: this.properties.listType ?? 'bulleted',
      listStyle: this.properties.listStyle ?? 'disc',
      title: (this.properties.title ?? '').trim(),
      titleColor: this.properties.titleColor, // ADD THIS LINE
      itemsColor: this.properties.itemsColor, // ADD THIS LINE
      columns: this.properties.columns ?? 1,
      minColumnWidth: this.properties.minColumnWidth ?? 150,
      themeVariant: this._themeVariant,
      debugHeadings: this.properties.debugHeadings ?? false,
      scope: this.properties.scope ?? 'page',
      groupBySection: this.properties.groupBySection ?? false,
      showSectionHeadings: this.properties.showSectionHeadings ?? false,
      titleLevel: this.properties.titleLevel ?? 'h2'
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    this.properties.scope ??= 'page';
    this.properties.groupBySection ??= false;
    this.properties.showSectionHeadings ??= false;
    this.properties.showH2 ??= true;
    this.properties.showH3 ??= true;
    this.properties.showH4 ??= true;
    this.properties.listType ??= 'bulleted';
    this.properties.listStyle ??= 'disc';
    this.properties.title ??= 'Contents';
    this.properties.columns ??= 1;
    this.properties.minColumnWidth ??= 150;
  this.properties.titleLevel ??= 'h2';
    this.properties.titleColorMode ??= 'theme';
  this.properties.itemsColorMode ??= 'theme';
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, (args: ThemeChangedEventArgs) => {
      this._themeVariant = args.theme;
      this.context.propertyPane.refresh();
      this.render();
    });
  }

protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    // We use unknown instead of any to satisfy ESLint. 
    // This method is called whenever a property pane field is updated.
    
    if (propertyPath === 'titleColorMode' || propertyPath === 'itemsColorMode' || propertyPath === 'listType' || propertyPath === 'groupBySection') {
      // Refreshing the pane allows conditional fields (like list styles) 
      // to update immediately based on the new selection.
      this.context.propertyPane.refresh();
    } 

    // Call the base class method to ensure standard behavior
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const p = this._themeVariant?.palette;
    /*
    const themeSwatches = [
      { color: p?.themePrimary ?? '#0078d4', label: 'Theme primary' },
      { color: p?.neutralPrimary ?? '#333333', label: 'Neutral primary' },
      { color: p?.white ?? '#ffffff', label: 'White' },
    ];
    */

    const themeSwatches = [
      { color: p?.themePrimary ?? '#0078d4', label: 'Theme primary' },
      { color: p?.themeSecondary ?? '#2b88d8', label: 'Theme secondary' },
      { color: p?.accent ?? '#ffb900', label: 'Accent' },
      { color: p?.neutralPrimary ?? '#333333', label: 'Neutral primary' },
      { color: p?.neutralSecondary ?? '#666666', label: 'Neutral secondary' },
      { color: p?.neutralDark ?? '#212121', label: 'Neutral dark' },
      { color: p?.black ?? '#000000', label: 'Black' },
      { color: p?.white ?? '#ffffff', label: 'White' },
    ];

    
    const standardSwatches = [
      { color: '#e81123', label: 'Red' },
      { color: '#ff8c00', label: 'Orange' },
      { color: '#ffb900', label: 'Yellow' },
      { color: '#107c10', label: 'Green' },
      { color: '#00bcf2', label: 'Light blue' },
      { color: '#5c2d91', label: 'Purple' },
      { color: '#605e5c', label: 'Gray' },
    ];



// Use <unknown> instead of <any> or <ITableOfContentsWebPartProps>
  // This is the standard way to hold a collection of different property pane controls
  const titleColorFields: IPropertyPaneField<unknown>[] = [];
  const itemColorFields: IPropertyPaneField<unknown>[] = [];

  // 3. ADD THE DROPDOWNS
  titleColorFields.push(PropertyPaneDropdown('titleColorMode', {
    label: 'Title Color Mode',
    options: [
      { key: 'theme', text: 'Theme Colours' },
      { key: 'standard', text: 'Standard Colours' },
      { key: 'custom', text: 'Custom Color' }
    ],
    selectedKey: this.properties.titleColorMode
  }));

  itemColorFields.push(PropertyPaneDropdown('itemsColorMode', {
    label: 'Items Color Mode',
    options: [
      { key: 'theme', text: 'Theme Colours' },
      { key: 'standard', text: 'Standard Colours' },
      { key: 'custom', text: 'Custom Color' }
    ],
    selectedKey: this.properties.itemsColorMode
  }));

  // 4. CONDITIONAL LOGIC FOR TITLE (Now titleColorFields is definitely defined)
  if (this.properties.titleColorMode === 'theme') {
    titleColorFields.push(PropertyFieldSwatchColorPicker('titleColor', {
      label: 'Select Theme Color',
      selectedColor: this.properties.titleColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      colors: themeSwatches,
      key: 'titleThemeColors'
    }));
  } else if (this.properties.titleColorMode === 'standard') {
    titleColorFields.push(PropertyFieldSwatchColorPicker('titleColor', {
      label: 'Select Standard Color',
      selectedColor: this.properties.titleColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      colors: standardSwatches,
      key: 'titleStdColors'
    }));
  } else {
    titleColorFields.push(PropertyFieldColorPicker('titleColor', {
      label: 'Choose Custom Color',
      selectedColor: this.properties.titleColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      style: PropertyFieldColorPickerStyle.Inline,
      key: 'titleCustomColor'
    }));
  }

  // 5. CONDITIONAL LOGIC FOR ITEMS
  if (this.properties.itemsColorMode === 'theme') {
    itemColorFields.push(PropertyFieldSwatchColorPicker('itemsColor', {
      label: 'Select Theme Color',
      selectedColor: this.properties.itemsColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      colors: themeSwatches,
      key: 'itemsThemeColors'
    }));
  } else if (this.properties.itemsColorMode === 'standard') {
    itemColorFields.push(PropertyFieldSwatchColorPicker('itemsColor', {
      label: 'Select Standard Color',
      selectedColor: this.properties.itemsColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      colors: standardSwatches,
      key: 'itemsStdColors'
    }));
  } else {
    itemColorFields.push(PropertyFieldColorPicker('itemsColor', {
      label: 'Choose Custom Color',
      selectedColor: this.properties.itemsColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      properties: this.properties,
      style: PropertyFieldColorPickerStyle.Inline,
      key: 'itemsCustomColor'
    }));
  }



    // FIX: Create a unique combined list to prevent rendering bugs
    /*
  const allSwatches = [...themeSwatches, ...standardSwatches].filter((v, i, a) => 
    a.findIndex(t => (t.color.toLowerCase() === v.color.toLowerCase())) === i
  );
  */

    /*
    */

    
    const bulletStyleOptions = [
      { key: 'disc', text: 'Bullet (●)' },
      { key: 'circle', text: 'Circle (○)' },
      { key: 'square', text: 'Square (■)' },
      { key: 'none', text: 'None' }
    ];

    const numberStyleOptions = [
      { key: 'decimal', text: 'Decimal (hierarchical 1., 1.1, 1.1.1 …)' },
      { key: 'lower-alpha', text: 'Lower-alpha (a, b, c, …)' },
      { key: 'upper-alpha', text: 'Upper-alpha (A, B, C, …)' },
      { key: 'lower-roman', text: 'Lower-roman (i, ii, iii, …)' },
      { key: 'upper-roman', text: 'Upper-roman (I, II, III, …)' }
    ];

    const styleOptions = this.properties.listType === 'numbered'
      ? numberStyleOptions
      : bulletStyleOptions;
    

    return {
      pages: [{
        header: { description: 'Table of Contents settings' },
        groups: [
          {
            groupName: 'Webpart Title & Colour',
            groupFields: [
              PropertyPaneTextField('title', { label: 'Title (optional)' }),
              ...titleColorFields,
                PropertyPaneDropdown('titleLevel', {
                  label: 'Title tag',
                  options: [
                    { key: 'h2', text: 'H2 (largest)' },
                    { key: 'h3', text: 'H3' },
                    { key: 'h4', text: 'H4 (smallest)' }
                  ],
                  selectedKey: this.properties.titleLevel ?? 'h2'
                })
            ]
          },
          {
            groupName: 'Heading levels',
            groupFields: [
              PropertyPaneCheckbox('showH2', { text: 'Include H2' }),
              PropertyPaneCheckbox('showH3', { text: 'Include H3' }),
              PropertyPaneCheckbox('showH4', { text: 'Include H4' })
            ]
          },
          {
            groupName: 'Formatting',
            groupFields: [
              PropertyPaneDropdown('listType', {
                label: 'List type',
                options: [{ key: 'bulleted', text: 'Bulleted' }, { key: 'numbered', text: 'Numbered' }],
                selectedKey: this.properties.listType ?? 'bulleted'
              }),
              PropertyPaneDropdown('listStyle', {
                label: 'Style',
                options: styleOptions,
                selectedKey: this.properties.listStyle ?? 'disc'

              }),
              ...itemColorFields // Spread the pre-built array

            ]
          },
          {
            groupName: 'Columns',
            groupFields: [
              PropertyPaneDropdown('columns', {
                label: 'Columns',
                options: [{ key: 1, text: '1' }, { key: 2, text: '2' }, { key: 3, text: '3' }]
              })
            ]
          },
          {
            groupName: 'Scope',
            groupFields: [
              PropertyPaneDropdown('scope', {
                label: 'Scan scope',
                options: [{ key: 'page', text: 'Whole page' }, { key: 'section', text: 'This section only' }]
              })
            ]
          },
          {
            groupName: 'Grouping',
            groupFields: [
              PropertyPaneCheckbox('groupBySection', { text: 'Group items by page section' }),
              PropertyPaneCheckbox('showSectionHeadings', { text: 'Show section headings' })
            ]
          }
        ]
      }]
    };
  }
}