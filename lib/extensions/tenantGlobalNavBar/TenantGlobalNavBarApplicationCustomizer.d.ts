import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITenantGlobalNavBarApplicationCustomizerProperties {
    TopMenuTermSet?: string;
    BottomMenuTermSet?: string;
}
/** A Custom Action which can be run during execution of a Client Side Application */
export default class TenantGlobalNavBarApplicationCustomizer extends BaseApplicationCustomizer<ITenantGlobalNavBarApplicationCustomizerProperties> {
    private _topPlaceholder;
    private _bottomPlaceholder;
    private _topMenuItems;
    private _bottomMenuItems;
    onInit(): Promise<void>;
    private _renderPlaceHolders();
    private _onDispose();
}
