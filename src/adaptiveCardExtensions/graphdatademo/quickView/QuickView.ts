import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'GraphdatademoAdaptiveCardExtensionStrings';
import { IGraphdatademoAdaptiveCardExtensionProps, IGraphdatademoAdaptiveCardExtensionState } from '../GraphdatademoAdaptiveCardExtension';

export interface IQuickViewData {
  message:any
}

export class QuickView extends BaseAdaptiveCardView<
  IGraphdatademoAdaptiveCardExtensionProps,
  IGraphdatademoAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      message: this.state.currentEmail
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}