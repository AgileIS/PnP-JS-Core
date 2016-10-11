"use strict";

import { HttpClientImpl, FetchOptions } from "./httpclient";
import { NTLM } from "./ntlm";
import { Util } from "../utils/util";
import * as http from "http";
import * as https from "https";

declare var global: any;
declare var require: (path: string) => any;
let nodeFetch = require("node-fetch");

/**
 * Fetch client for use within nodejs, using ntlm auth
 */
export class NodeHttpNtlmClient implements HttpClientImpl {

    private _agent: http.Agent | https.Agent;

    constructor(private _username: string, private _password: string, private _workstation: string, private _domain: string) {
        // here we "cheat" and set the globals for fetch things when this client is instantiated
        global.Headers = nodeFetch.Headers;
        global.Request = nodeFetch.Request;
        global.Response = nodeFetch.Response;
    }

    public fetch(url: string, options?: FetchOptions): Promise<Response> {
        // maxSockets = 1, because handshake error offer multiple sockets
        // (sometimes second handshake call goes out before agent frees the socket => new socket for second call)
        // ToDo: inserting global agent for one time auth (all calls offer 1 socket)
        if (!this._agent && url.indexOf("https") > -1) {
            this._agent = new https.Agent({ keepAlive: true, maxSockets: 1 });
        } else if (!this._agent) {
            this._agent = new http.Agent({ keepAlive: true, maxSockets: 1 });
        }
        let newHeader = new Headers();
        newHeader.append("Connection", "keep-alive");
        this._mergeHeaders(newHeader, options.headers);
        let extendedOptions = Util.extend(options, { agent: this._agent,  headers: newHeader }, false);

        /*return nodeFetch(url, extendedOptions).then((response: Response) => {
            if (response.status === 401) {
                return this.fetchWithHandshake(url, extendedOptions);
            }
            return new Promise<Response>(resolve => resolve(response));
        });*/
        return this.fetchWithHandshake(url, extendedOptions);
    }

    private fetchWithHandshake(url: string, options?: any): Promise<Response> {

        let ntlmOptions = {
            domain: this._domain,
            password: this._password,
            username: this._username,
            workstation: this._workstation,
        };
        options.headers.set("Authorization", NTLM.createType1Message(ntlmOptions));

        return nodeFetch(url, options).then((response: Response) => {
            let type2msg = NTLM.parseType2Message(response.headers.get("www-authenticate"), (error: Error) => {
                return new Promise<Response>(reject => reject(response));
            });
            let type3msg = NTLM.createType3Message(type2msg, ntlmOptions);
            options.headers.set("Authorization", type3msg);
            return nodeFetch(url, options);
        });
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
