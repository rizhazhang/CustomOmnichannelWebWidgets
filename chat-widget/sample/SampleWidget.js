import * as React from "react";

import { BroadcastService } from "../lib/esm/index.js";
import LiveChatWidget from "../lib/esm/components/livechatwidget/LiveChatWidget.js";
import { OmnichannelChatSDK } from "@microsoft/omnichannel-chat-sdk";
import ReactDOM from "react-dom";
import { getCustomizationJson } from "./getCustomizationJson";
import { registerCacheWidgetStateEvent, restoreWidgetStateIfExistInCache } from "./cacheWidgetState.js";
import { getUnreadMessageCount } from "./getUnreadMessageCount";
import { clientDataStoreProvider } from "./Common/clientDataStoreProvider";
import { memoryDataStore } from "./Common/MemoryDataStore";
import * as microsoftTeams from "@microsoft/teams-js";

/* eslint @typescript-eslint/no-explicit-any: "off" */

let liveChatWidgetProps;

const main = async () => {
    console.info("main,starting...");
    const queryString = window.location.search;
    const urlParams = new URLSearchParams(queryString);
    const orgId = urlParams.get("data-org-id");
    const orgUrl = urlParams.get("data-org-url");
    const appId = urlParams.get("data-app-id");

    const script = document.getElementById("oc-lcw-script");
    const omnichannelConfig = {
        orgId: "4bdf7a19-29f1-4744-a45f-d041227058fe",//orgId ?? script?.getAttribute("data-org-id"),
        orgUrl: "https://unq4bdf7a1929f14744a45fd04122705-crm5.omnichannelengagementhub.com",//orgUrl ?? script?.getAttribute("data-org-url"),
        widgetId: "12707d66-6cfa-4c37-bd8e-7ed0244667c8"//appId ?? script?.getAttribute("data-app-id")
    };

    ////===========================
    const chatSDKConfig = {
        persistentChat: {
            disable: false,
            tokenUpdateTime: 21600000
        },
        getAuthToken: async () => {
            //const response = await fetch("http://contosohelp.com/token");
            var response = GetTeamsToken();//microsoftTeams.authentication.getAuthToken();
            console.info("Got auth response.");            
            // if (response) {
            //     return response;
            // }
            // else {
            //     return null;
            // }
            return response;
        }
    };

    const chatSDK = new OmnichannelChatSDK(omnichannelConfig);
    await chatSDK.initialize();
    
    // chatSDK.setContextProvider(function contextProvider(){
    //     //Here it is assumed that the corresponding work stream would have context variables with logical name of 'contextKey1', 'contextKey2', 'contextKey3'. If no context variable exists with a matching logical name, items are created assuming Type:string               
    //     return {"Account": {"Email": "vkuser1@teamsshows.com"}};
    // });
    
    console.info("initialized.");
    const chatConfig = await chatSDK.getLiveChatConfig();
    await registerCacheWidgetStateEvent();
    memoryDataStore();
    const widgetStateFromCache = await restoreWidgetStateIfExistInCache();
    console.info(widgetStateFromCache);
    const widgetStateJson = widgetStateFromCache ? JSON.parse(widgetStateFromCache) : undefined;
    console.info("widgetStateJson");
    await getUnreadMessageCount();
    const switchConfig = (config) => {
        liveChatWidgetProps = config;
        liveChatWidgetProps = {
            ...liveChatWidgetProps,
            chatSDK: chatSDK,
            chatConfig: chatConfig,
            liveChatContextFromCache: widgetStateJson,
            contextDataStore: clientDataStoreProvider(),
            controlProps: {
                skipChatButtonRendering: true
            }
        };

        ReactDOM.render(
            <LiveChatWidget {...liveChatWidgetProps} />,
            document.getElementById("oc-lcw-container")
        );
    };
    const startProactiveChat = (notificationUIConfig, showPrechat, inNewWindow) => {
        const startProactiveChatEvent = {
            eventName: "StartProactiveChat",
            payload: {
                bodyTitle: (notificationUIConfig && notificationUIConfig.message) ? notificationUIConfig.message : "Hello Customer",
                showPrechat: showPrechat,
                inNewWindow: inNewWindow
            }
        };
        BroadcastService.postMessage(startProactiveChatEvent);
    };

    window["switchConfig"] = switchConfig;
    window["startProactiveChat"] = startProactiveChat;
    switchConfig(await getCustomizationJson());
};


function GetTeamsToken()
{
    // const myObject = 
    // {
    //     "name":"john doe",
    //     "age": 32,
    //     "gender" : "male",
    //     "profession" : "optician" 
    // };
      
    // window.localStorage.setItem("myObject", JSON.stringify(myObject));
    // let newObject = window.localStorage.getItem("myObject");
    // console.info(JSON.parse(newObject));

    // microsoftTeams.app.initialize();
    
    try{
        display("1. Get auth token from Microsoft Teams");
        microsoftTeams.app.initialize().then(()=>{
            microsoftTeams.authentication.getAuthToken().then((result) => {
                display("2. Got auth token from Microsoft Teams");
                console.info(result);
                return result;
            }).catch((error) => {
                console.info(error);
            });

            console.info(microsoftTeams.app.getContext());
            

        });
    }catch (e) {
        console.info(e);
    }
    return null;
    // Get the user context from Teams and set it in the state
    // microsoftTeams.app.getContext((context, error) => {
    //     console.info("getContext");
    //     console.info(context.user);
    //     console.info(error.message);
    // }); 

    // microsoftTeams.app.getFrameContext()((context, error) => {
    //     console.info("getFrameContext");
    //     console.info(context.user);
    //     console.info(error.message);
    // }); 
    
    // microsoftTeams.authentication.getAuthToken().then((result) => {
    //     display("2. Got auth token from Microsoft Teams");
    //     display(result);
    //     return result;
    // }).catch((error) => {
    //     display("3. Failed to get auth token from Microsoft Teams");
    //     display(error.message);
    //     display(error);
    //     return ("Error getting token: " + error);
    // });
           
   
}


function getClientSideToken() {

    return new Promise((resolve, reject) => {
        display("1. Get auth token from Microsoft Teams");
        microsoftTeams.authentication.getAuthToken().then((result) => {
            display(result);
            resolve(result);
        }).catch((error) => {
            reject("Error getting token: " + error);
        });
    });
}

// 2. Exchange that token for a token with the required permissions
//    using the web service (see /auth/token handler in app.js)
function getServerSideToken(clientSideToken) {
    return new Promise((resolve, reject) => {
        microsoftTeams.app.getContext().then((context) => {
            fetch("/getProfileOnBehalfOf", {
                method: "post",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    "tid": context.user.tenant.id,
                    "token": clientSideToken
                }),
                mode: "cors",
                cache: "default"
            }).then((response) => {
                if (response.ok) {
                    return response.json();
                } else {
                    reject(response.error);
                }
            }).then((responseJson) => {
                if (responseJson.error) {
                    reject(responseJson.error);
                } else {
                    const profile = responseJson;
                    resolve(profile);
                }
            });
        });
    });
}

// 3. Get the server side token and use it to call the Graph API
function useServerSideToken(data) {

    display("2. Call https://graph.microsoft.com/v1.0/me/ with the server side token");
    return display(JSON.stringify(data, undefined, 4), "pre");
}

// Show the consent pop-up
function requestConsent() {
    return new Promise((resolve, reject) => {
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/auth-start",
            width: 600,
            height: 535}).then((result) => {
            let data = localStorage.getItem(result);
            localStorage.removeItem(result);
            resolve(data);
        }).catch((reason) => {
            reject(JSON.stringify(reason));
        });
    });
}

// Add text to the display in a <p> or other HTML element
function display(text, elementTag) {
    // var logDiv = document.getElementById("logs");
    // var p = document.createElement(elementTag ? elementTag : "p");
    // p.innerText = text;
    // logDiv.append(p);
    // console.log("ssoDemo: " + text);
    // return p;
    console.info("sso log:" + text);
}

// // In-line code
// await getClientSideToken()
//     .then((clientSideToken) => {
//         return getServerSideToken(clientSideToken);
//     }).then((profile) => {
//         return useServerSideToken(profile);
//     }).catch((error) => {
//         if (error === "invalid_grant") {
//             display(`Error: ${error} - user or admin consent required`);           
           
//         } else {
//             // Something else went wrong
//             display(`Error from web service: ${error}`);
//         }
//     });

GetTeamsToken();

main();


