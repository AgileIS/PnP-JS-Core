"use strict";

import { FetchOptions, HttpClientImpl } from "./httpclient";
import { Util } from "../utils/util";

declare var global: any;
declare var require: (path: string) => any;
let nodeFetch = require("node-fetch");

export class NodeHttpBasicClient implements HttpClientImpl {
    private authKey: string = "Authorization";
    private authValue: string;

    constructor(username: string, password: string) {
        global.Headers = nodeFetch.Headers;
        global.Request = nodeFetch.Request;
        global.Response = nodeFetch.Response;

        this.authValue = `Basic ${new Buffer(`${username}:${password}`).toString("base64")}`;
    }

    public fetch(url: string, options?: FetchOptions): Promise<Response> {
        let newHeader = new Headers();
        newHeader.append(this.authKey, this.authValue);
        this._mergeHeaders(newHeader, options.headers);
        let extendedOptions = Util.extend(options, { headers: newHeader }, false);
        return nodeFetch(url, extendedOptions);
    }

    private _mergeHeaders(target: Headers, source: any): void {
        if (typeof source !== "undefined" && source !== null) {
            let temp = <any> new Request("", { headers: source });
            temp.headers.forEach((value: string, name: string) => {
                if (name.toLowerCase() === "accept" && value.toLowerCase() === "application/json") {
                    target.append(name, "application/json;odata=verbose");
                } else {
                    target.append(name, value);
                }
            });
        }
    }
}


