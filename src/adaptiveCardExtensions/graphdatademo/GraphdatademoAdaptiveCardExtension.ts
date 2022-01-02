import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { GraphdatademoPropertyPane } from './GraphdatademoPropertyPane';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphdatademoAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
}

export interface IGraphdatademoAdaptiveCardExtensionState {
  emails: any;
  currentIndex:any;
  currentEmail:any;
}

const CARD_VIEW_REGISTRY_ID: string = 'Graphdatademo_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Graphdatademo_QUICK_VIEW';

export default class GraphdatademoAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IGraphdatademoAdaptiveCardExtensionProps,
  IGraphdatademoAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: GraphdatademoPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      emails: [],
      currentIndex:0,
      currentEmail:{}
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    this.getOutlookData()

    return Promise.resolve();
  }

  private getOutlookData(){
    
    this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api("/me/mailfolders/Inbox/messages").get((error, messages: any) => {
        console.log(messages);
        this.setState({currentEmail:messages.value[this.state.currentIndex],emails:messages.value});
      });
    });
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Graphdatademo-property-pane'*/
      './GraphdatademoPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.GraphdatademoPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
