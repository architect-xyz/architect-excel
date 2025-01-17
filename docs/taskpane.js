(()=>{"use strict";var e={101:function(e,t,n){var r=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(o,i){function a(e){try{u(r.next(e))}catch(e){i(e)}}function c(e){try{u(r.throw(e))}catch(e){i(e)}}function u(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,c)}u((r=r.apply(e,t||[])).next())}))},o=this&&this.__generator||function(e,t){var n,r,o,i={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]},a=Object.create(("function"==typeof Iterator?Iterator:Object).prototype);return a.next=c(0),a.throw=c(1),a.return=c(2),"function"==typeof Symbol&&(a[Symbol.iterator]=function(){return this}),a;function c(c){return function(u){return function(c){if(n)throw new TypeError("Generator is already executing.");for(;a&&(a=0,c[0]&&(i=0)),i;)try{if(n=1,r&&(o=2&c[0]?r.return:c[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,c[1])).done)return o;switch(r=0,o&&(c=[2&c[0],o.value]),c[0]){case 0:case 1:o=c;break;case 4:return i.label++,{value:c[1],done:!1};case 5:i.label++,r=c[1],c=[0];continue;case 7:c=i.ops.pop(),i.trys.pop();continue;default:if(!((o=(o=i.trys).length>0&&o[o.length-1])||6!==c[0]&&2!==c[0])){i=0;continue}if(3===c[0]&&(!o||c[1]>o[0]&&c[1]<o[3])){i.label=c[1];break}if(6===c[0]&&i.label<o[1]){i.label=o[1],o=c;break}if(o&&i.label<o[2]){i.label=o[2],i.ops.push(c);break}o[2]&&i.ops.pop(),i.trys.pop();continue}c=t.call(e,i)}catch(e){c=[6,e],r=0}finally{n=o=0}if(5&c[0])throw c[1];return{value:c[0]?c[1]:void 0,done:!0}}([c,u])}}};Object.defineProperty(t,"__esModule",{value:!0}),t.fetchData=function(){return r(this,void 0,void 0,(function(){var e,t;return o(this,(function(n){if(e=Office.context.document.settings.get("apiKey"),t=Office.context.document.settings.get("apiSecret"),!e||!t)throw new Error("API key or secret not set. Use the ribbon to configure them.");return[2,"Using API Key: ".concat(e,", Secret: ").concat(t)]}))}))},t.initializeClient=function(){var e=Office.context.document.settings.get("apiKey"),t=Office.context.document.settings.get("apiSecret");if(!e||!t)throw new Error("API Key and Secret must be provided.");c.apiKey=e,c.apiSecret=t,i=(0,a.create)(c)},t.fetchMarketSnapshot=function(e){return r(this,void 0,void 0,(function(){var t,n,r,a;return o(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,i.marketSnapshot([],e)];case 1:return(t=o.sent())&&t.bidPrice&&t.askPrice?(n=parseFloat(t.bidPrice),r=parseFloat(t.askPrice),[2,isNaN(n)||isNaN(r)?NaN:(n+r)/2]):(console.error("Invalid or missing snapshot data"),[2,NaN]);case 2:return a=o.sent(),console.error("Error fetching market snapshot:",a),[2,void 0];case 3:return[2]}}))}))};var i,a=n(273),c={host:"https://app.architect.co",apiKey:"",apiSecret:"",tradingMode:"live"};Office.onReady((function(e){e.host===Office.HostType.Excel&&console.log("Excel Add-in ready.")}))},905:function(e,t,n){var r=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(o,i){function a(e){try{u(r.next(e))}catch(e){i(e)}}function c(e){try{u(r.throw(e))}catch(e){i(e)}}function u(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,c)}u((r=r.apply(e,t||[])).next())}))},o=this&&this.__generator||function(e,t){var n,r,o,i={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]},a=Object.create(("function"==typeof Iterator?Iterator:Object).prototype);return a.next=c(0),a.throw=c(1),a.return=c(2),"function"==typeof Symbol&&(a[Symbol.iterator]=function(){return this}),a;function c(c){return function(u){return function(c){if(n)throw new TypeError("Generator is already executing.");for(;a&&(a=0,c[0]&&(i=0)),i;)try{if(n=1,r&&(o=2&c[0]?r.return:c[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,c[1])).done)return o;switch(r=0,o&&(c=[2&c[0],o.value]),c[0]){case 0:case 1:o=c;break;case 4:return i.label++,{value:c[1],done:!1};case 5:i.label++,r=c[1],c=[0];continue;case 7:c=i.ops.pop(),i.trys.pop();continue;default:if(!((o=(o=i.trys).length>0&&o[o.length-1])||6!==c[0]&&2!==c[0])){i=0;continue}if(3===c[0]&&(!o||c[1]>o[0]&&c[1]<o[3])){i.label=c[1];break}if(6===c[0]&&i.label<o[1]){i.label=o[1],o=c;break}if(o&&i.label<o[2]){i.label=o[2],i.ops.push(c);break}o[2]&&i.ops.pop(),i.trys.pop();continue}c=t.call(e,i)}catch(e){c=[6,e],r=0}finally{n=o=0}if(5&c[0])throw c[1];return{value:c[0]?c[1]:void 0,done:!0}}([c,u])}}};Object.defineProperty(t,"__esModule",{value:!0});var i=n(101);document.addEventListener("DOMContentLoaded",(function(){var e=document.getElementById("api-form");null==e||e.addEventListener("submit",(function(e){return r(void 0,void 0,void 0,(function(){var t,n,r,a,c;return o(this,(function(o){if(e.preventDefault(),t=null===(a=document.getElementById("apiKey"))||void 0===a?void 0:a.value.trim(),n=null===(c=document.getElementById("apiSecret"))||void 0===c?void 0:c.value.trim(),r=document.getElementById("status"),!t||!n)return r.textContent="API Key and Secret are required.",[2];try{Office.context.document.settings.set("apiKey",t),Office.context.document.settings.set("apiSecret",n),Office.context.document.settings.saveAsync(),r.textContent="Credentials saved!",(0,i.initializeClient)()}catch(e){r.textContent="Error: ".concat(e.message)}return[2]}))}))}))}))},273:(e,t,n)=>{n.r(t);var r=Object.create?function(e,t,n,r){void 0===r&&(r=n);var o=Object.getOwnPropertyDescriptor(t,n);o&&!("get"in o?!t.__esModule:o.writable||o.configurable)||(o={enumerable:!0,get:function(){return t[n]}}),Object.defineProperty(e,r,o)}:function(e,t,n,r){void 0===r&&(r=n),e[r]=t[n]},o=Object.create?function(e,t){Object.defineProperty(e,"default",{enumerable:!0,value:t})}:function(e,t){e.default=t};Object.defineProperty(exports,"__esModule",{value:!0}),exports.enums=exports.L1BookSnapshot=void 0,exports.create=function(e){return new a.Client(e,i.createClient)};const i=require("graphql-http");exports.L1BookSnapshot=function(e){if(e&&e.__esModule)return e;var t={};if(null!=e)for(var n in e)"default"!==n&&Object.prototype.hasOwnProperty.call(e,n)&&r(t,e,n);return o(t,e),t}(require("./grpc/l1booksnapshot.js"));const a=require("./sdk.js"),c=require("./graphql/graphql.js");exports.enums={AlgoControlCommand:c.AlgoControlCommand,AlgoKind:c.AlgoKind,AlgoRunningStatus:c.AlgoRunningStatus,CandleWidth:c.CandleWidth,CmeSecurityType:c.CmeSecurityType,CreateOrderType:c.CreateOrderType,CreateTimeInForceInstruction:c.CreateTimeInForceInstruction,EnvironmentKind:c.EnvironmentKind,FillKind:c.FillKind,LicenseTier:c.LicenseTier,MMAlgoKind:c.MMAlgoKind,MinOrderQuantityUnit:c.MinOrderQuantityUnit,OrderSource:c.OrderSource,OrderStateFlags:c.OrderStateFlags,Reason:c.Reason,ReferencePrice:c.ReferencePrice}}},t={};function n(r){var o=t[r];if(void 0!==o)return o.exports;var i=t[r]={exports:{}};return e[r].call(i.exports,i,i.exports,n),i.exports}n.r=e=>{"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n(905)})();