import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { escape } from '@microsoft/sp-lodash-subset';
//import styles from './StockHeatmapWidgetWebPart.module.scss';
import * as strings from 'StockHeatmapWidgetWebPartStrings';

export interface IStockHeatmapWidgetWebPartProps {
  description: string;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare const TradingView: any;

export default class StockHeatmapWidgetWebPart extends BaseClientSideWebPart<IStockHeatmapWidgetWebPartProps> {

  //private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // 输出 Web Part 的 HTML 内容，包含 TradingView widget 的容器
    this.domElement.innerHTML = `
      <div>
        <h2>Welcome, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>Environment: ${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>TradingView Stock Heatmap</h3>
        <div class="tradingview-widget-container">
          <div class="tradingview-widget-container__widget" id="tradingview_stockheatmap" style="width:100%; height:600px;"></div>
          <div class="tradingview-widget-copyright">
            <a href="https://www.tradingview.com/" rel="noopener nofollow" target="_blank">
              <span class="blue-text">Track all markets on TradingView</span>
            </a>
          </div>
        </div>
      </div>`;

    // 动态加载 TradingView 脚本并初始化 widget
    SPComponentLoader.loadScript("https://s3.tradingview.com/external-embedding/embed-widget-stock-heatmap.js", { globalExportsName: "TradingView" })
      .then(() => {
        const TVWidget = new TradingView.widget({
          "container_id": "tradingview_stockheatmap",
          "width": "100%",
          "height": "100%",
          "exchanges": [],
          "dataSource": "SPX500",
          "grouping": "sector",
          "blockSize": "market_cap_basic",
          "blockColor": "change",
          "locale": "en",
          "symbolUrl": "https://www.tradingview.com/symbols",
          "colorTheme": "dark",
          "hasTopBar": true,
          "isDataSetEnabled": true,
          "isZoomEnabled": true,
          "hasSymbolTooltip": true,
          "isMonoSize": false
        });
        console.log("TradingView Widget initialized:", TVWidget);
      })
      .catch(error => {
        console.error("加载 TradingView 脚本失败：", error);
      });
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  /* protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  } */

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
