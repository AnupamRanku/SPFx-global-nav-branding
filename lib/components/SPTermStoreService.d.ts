import { SPHttpClient } from '@microsoft/sp-http';
/**
 * @interface
 * Interface for SPTermStoreService configuration
 */
export interface ISPTermStoreServiceConfiguration {
    spHttpClient: SPHttpClient;
    siteAbsoluteUrl: string;
}
/**
 * @interface
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
    identity: string;
    isAvailableForTagging: boolean;
    name: string;
    guid: string;
    customSortOrder: string;
    terms: ISPTermObject[];
    localCustomProperties: any;
}
/**
 * @class
 * Service implementation to manage term stores in SharePoint
 * Basic implementation taken from: https://oliviercc.github.io/sp-client-custom-fields/
 */
export declare class SPTermStoreService {
    private spHttpClient;
    private siteAbsoluteUrl;
    private formDigest;
    /**
     * @function
     * Service constructor
     */
    constructor(config: ISPTermStoreServiceConfiguration);
    /**
     * @function
     * Gets the collection of term stores in the current SharePoint env
     */
    getTermsFromTermSetAsync(termSetName: string): Promise<ISPTermObject[]>;
    /**
     * @function
     * Gets the child terms of another term of the Term Store in the current SharePoint env
     */
    private getChildTermsAsync(term);
    /**
     * @function
     * Projects a Term object into an object of type ISPTermObject, including child terms
     * @param guid
     */
    private projectTermAsync(term);
    /**
     * @function
     * Clean the Guid from the Web Service response
     * @param guid
     */
    private cleanGuid(guid);
}
