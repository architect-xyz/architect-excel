!function(e,n){"object"==typeof exports&&"object"==typeof module?module.exports=n():"function"==typeof define&&define.amd?define([],n):"object"==typeof exports?exports.functions=n():e.functions=n()}(this,(()=>(()=>{"use strict";var e={985:(e,n,t)=>{t.d(n,{$E:()=>r,$W:()=>i});let i={host:"https://app.architect.co/",apiKey:"",apiSecret:"",tradingMode:"live"};async function r(e){if("undefined"!=typeof Office&&Office.context&&"undefined"!=typeof OfficeRuntime)return await OfficeRuntime.storage.getItem(e);if("undefined"!=typeof localStorage)return localStorage.getItem(e);throw new Error("No available storage method to get from.")}}},n={};function t(i){var r=n[i];if(void 0!==r)return r.exports;var o=n[i]={exports:{}};return e[i](o,o.exports,t),o.exports}t.d=(e,n)=>{for(var i in n)t.o(n,i)&&!t.o(e,i)&&Object.defineProperty(e,i,{enumerable:!0,get:n[i]})},t.o=(e,n)=>Object.prototype.hasOwnProperty.call(e,n),t.r=e=>{"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})};var i={};function r(e){const{credentials:n="same-origin",referrer:t,referrerPolicy:i,shouldRetry:r=()=>!1}=e,a=e.fetchFn||fetch,s=e.abortControllerImpl||AbortController,c=(()=>{let e=!1;const n=[];return{get disposed(){return e},onDispose:t=>e?(setTimeout((()=>t()),0),()=>{}):(n.push(t),()=>{n.splice(n.indexOf(t),1)}),dispose(){if(!e){e=!0;for(const e of[...n])e()}}}})();return{subscribe(u,l){if(c.disposed)throw new Error("Client has been disposed");const d=new s,m=c.onDispose((()=>{m(),d.abort()}));return(async()=>{var s;let c=null,m=0;for(;;){if(c){const e=await r(c,m);if(d.signal.aborted)return;if(!e)throw c;m++}try{const r="function"==typeof e.url?await e.url(u):e.url;if(d.signal.aborted)return;const c="function"==typeof e.headers?await e.headers():null!==(s=e.headers)&&void 0!==s?s:{};if(d.signal.aborted)return;let m;try{m=await a(r,{signal:d.signal,method:"POST",headers:Object.assign(Object.assign({},c),{"content-type":"application/json; charset=utf-8",accept:"application/graphql-response+json, application/json"}),credentials:n,referrer:t,referrerPolicy:i,body:JSON.stringify(u)})}catch(e){throw new o(e)}if(!m.ok)throw new o(m);if(!m.body)throw new Error("Missing response body");const y=m.headers.get("content-type");if(!y)throw new Error("Missing response content-type");if(!y.includes("application/graphql-response+json")&&!y.includes("application/json"))throw new Error(`Unsupported response content-type ${y}`);const f=await m.json();return l.next(f),d.abort()}catch(e){if(d.signal.aborted)return;if(!(e instanceof o))throw e;c=e}}})().then((()=>l.complete())).catch((e=>l.error(e))),()=>d.abort()},dispose(){c.dispose()}}}t.r(i),t.d(i,{getMarketBBO:()=>ae,getMarketLast:()=>oe,getMarketMid:()=>se,initializeClient:()=>re,searchSymbols:()=>ce,testClient:()=>ue,testClient2:()=>le});class o extends Error{constructor(e){let n,t;var i;!function(e){return"object"==typeof e&&null!==e}(i=e)||"boolean"!=typeof i.ok||"number"!=typeof i.status||"string"!=typeof i.statusText?n=e instanceof Error?e.message:String(e):(t=e,n="Server responded with "+e.status+": "+e.statusText),super(n),this.name=this.constructor.name,this.response=t}}var a,s,c="Document",u="FragmentDefinition";class l extends Error{constructor(e,n,t,i,r,o,a){super(e),this.name="GraphQLError",this.message=e,r&&(this.path=r),n&&(this.nodes=Array.isArray(n)?n:[n]),t&&(this.source=t),i&&(this.positions=i),o&&(this.originalError=o);var s=a;if(!s&&o){var c=o.extensions;c&&"object"==typeof c&&(s=c)}this.extensions=s||{}}toJSON(){return{...this,message:this.message}}toString(){return this.message}get[Symbol.toStringTag](){return"GraphQLError"}}function d(e){return new l(`Syntax Error: Unexpected token at ${s} in ${e}`)}function m(e){if(e.lastIndex=s,e.test(a))return a.slice(s,s=e.lastIndex)}var y=/ +(?=[^\s])/y;function f(e){for(var n=e.split("\n"),t="",i=0,r=0,o=n.length-1,a=0;a<n.length;a++)y.lastIndex=0,y.test(n[a])&&(a&&(!i||y.lastIndex<i)&&(i=y.lastIndex),r=r||a,o=a);for(var s=r;s<=o;s++)s!==r&&(t+="\n"),t+=n[s].slice(i).replace(/\\"""/g,'"""');return t}function p(){for(var e=0|a.charCodeAt(s++);9===e||10===e||13===e||32===e||35===e||44===e||65279===e;e=0|a.charCodeAt(s++))if(35===e)for(;10!==(e=a.charCodeAt(s++))&&13!==e;);s--}function v(){for(var e=s,n=0|a.charCodeAt(s++);n>=48&&n<=57||n>=65&&n<=90||95===n||n>=97&&n<=122;n=0|a.charCodeAt(s++));if(e===s-1)throw d("Name");var t=a.slice(e,--s);return p(),t}function h(){return{kind:"Name",value:v()}}var g=/(?:"""|(?:[\s\S]*?[^\\])""")/y,b=/(?:(?:\.\d+)?[eE][+-]?\d+|\.\d+)/y;function $(e){var n;switch(a.charCodeAt(s)){case 91:s++,p();for(var t=[];93!==a.charCodeAt(s);)t.push($(e));return s++,p(),{kind:"ListValue",values:t};case 123:s++,p();for(var i=[];125!==a.charCodeAt(s);){var r=h();if(58!==a.charCodeAt(s++))throw d("ObjectField");p(),i.push({kind:"ObjectField",name:r,value:$(e)})}return s++,p(),{kind:"ObjectValue",fields:i};case 36:if(e)throw d("Variable");return s++,{kind:"Variable",name:h()};case 34:if(34===a.charCodeAt(s+1)&&34===a.charCodeAt(s+2)){if(s+=3,null==(n=m(g)))throw d("StringValue");return p(),{kind:"StringValue",value:f(n.slice(0,-3)),block:!0}}var o,c=s;s++;var u=!1;for(o=0|a.charCodeAt(s++);92===o&&(s++,u=!0)||10!==o&&13!==o&&34!==o&&o;o=0|a.charCodeAt(s++));if(34!==o)throw d("StringValue");return n=a.slice(c,s),p(),{kind:"StringValue",value:u?JSON.parse(n):n.slice(1,-1),block:!1};case 45:case 48:case 49:case 50:case 51:case 52:case 53:case 54:case 55:case 56:case 57:for(var l,y=s++;(l=0|a.charCodeAt(s++))>=48&&l<=57;);var I=a.slice(y,--s);if(46===(l=a.charCodeAt(s))||69===l||101===l){if(null==(n=m(b)))throw d("FloatValue");return p(),{kind:"FloatValue",value:I+n}}return p(),{kind:"IntValue",value:I};case 110:if(117===a.charCodeAt(s+1)&&108===a.charCodeAt(s+2)&&108===a.charCodeAt(s+3))return s+=4,p(),{kind:"NullValue"};break;case 116:if(114===a.charCodeAt(s+1)&&117===a.charCodeAt(s+2)&&101===a.charCodeAt(s+3))return s+=4,p(),{kind:"BooleanValue",value:!0};break;case 102:if(97===a.charCodeAt(s+1)&&108===a.charCodeAt(s+2)&&115===a.charCodeAt(s+3)&&101===a.charCodeAt(s+4))return s+=5,p(),{kind:"BooleanValue",value:!1}}return{kind:"EnumValue",value:v()}}function I(e){if(40===a.charCodeAt(s)){var n=[];s++,p();do{var t=h();if(58!==a.charCodeAt(s++))throw d("Argument");p(),n.push({kind:"Argument",name:t,value:$(e)})}while(41!==a.charCodeAt(s));return s++,p(),n}}function E(e){if(64===a.charCodeAt(s)){var n=[];do{s++,n.push({kind:"Directive",name:h(),arguments:I(e)})}while(64===a.charCodeAt(s));return n}}function S(){for(var e=0;91===a.charCodeAt(s);)e++,s++,p();var n={kind:"NamedType",name:h()};do{if(33===a.charCodeAt(s)&&(s++,p(),n={kind:"NonNullType",type:n}),e){if(93!==a.charCodeAt(s++))throw d("NamedType");p(),n={kind:"ListType",type:n}}}while(e--);return n}function T(){if(123!==a.charCodeAt(s++))throw d("SelectionSet");return p(),x()}function x(){var e=[];do{if(46===a.charCodeAt(s)){if(46!==a.charCodeAt(++s)||46!==a.charCodeAt(++s))throw d("SelectionSet");switch(s++,p(),a.charCodeAt(s)){case 64:e.push({kind:"InlineFragment",typeCondition:void 0,directives:E(!1),selectionSet:T()});break;case 111:110===a.charCodeAt(s+1)?(s+=2,p(),e.push({kind:"InlineFragment",typeCondition:{kind:"NamedType",name:h()},directives:E(!1),selectionSet:T()})):e.push({kind:"FragmentSpread",name:h(),directives:E(!1)});break;case 123:s++,p(),e.push({kind:"InlineFragment",typeCondition:void 0,directives:void 0,selectionSet:x()});break;default:e.push({kind:"FragmentSpread",name:h(),directives:E(!1)})}}else{var n=h(),t=void 0;58===a.charCodeAt(s)&&(s++,p(),t=n,n=h());var i=I(!1),r=E(!1),o=void 0;123===a.charCodeAt(s)&&(s++,p(),o=x()),e.push({kind:"Field",alias:t,name:n,arguments:i,directives:r,selectionSet:o})}}while(125!==a.charCodeAt(s));return s++,p(),{kind:"SelectionSet",selections:e}}function A(){if(p(),40===a.charCodeAt(s)){var e=[];s++,p();do{if(36!==a.charCodeAt(s++))throw d("Variable");var n=h();if(58!==a.charCodeAt(s++))throw d("VariableDefinition");p();var t=S(),i=void 0;61===a.charCodeAt(s)&&(s++,p(),i=$(!0)),p(),e.push({kind:"VariableDefinition",variable:{kind:"Variable",name:n},type:t,defaultValue:i,directives:E(!0)})}while(41!==a.charCodeAt(s));return s++,p(),e}}function O(){var e=h();if(111!==a.charCodeAt(s++)||110!==a.charCodeAt(s++))throw d("FragmentDefinition");return p(),{kind:"FragmentDefinition",name:e,typeCondition:{kind:"NamedType",name:h()},directives:E(!1),selectionSet:T()}}function C(){var e=[];do{if(123===a.charCodeAt(s))s++,p(),e.push({kind:"OperationDefinition",operation:"query",name:void 0,variableDefinitions:void 0,directives:void 0,selectionSet:x()});else{var n=v();switch(n){case"fragment":e.push(O());break;case"query":case"mutation":case"subscription":var t,i=void 0;40!==(t=a.charCodeAt(s))&&64!==t&&123!==t&&(i=h()),e.push({kind:"OperationDefinition",operation:n,name:i,variableDefinitions:A(),directives:E(!1),selectionSet:T()});break;default:throw d("Document")}}}while(s<a.length);return e}function D(e,n){return a=e.body?e.body:e,s=0,p(),n&&n.noLocation?{kind:"Document",definitions:C()}:{kind:"Document",definitions:C(),loc:{start:0,end:a.length,startToken:void 0,endToken:void 0,source:{body:a,name:"graphql.web",locationOffset:{line:1,column:1}}}}}var N=0,k=new Set;function V(){function e(e,n){var t,i,r=D(e).definitions,o=new Set;for(var a of n||[])for(var s of a.definitions)s.kind!==u||o.has(s)||(r.push(s),o.add(s));return(t=r[0].kind===u)&&r[0].directives&&(r[0].directives=r[0].directives.filter((e=>"_unmask"!==e.name.value))),{kind:c,definitions:r,get loc(){if(!i&&t){var r=e+function(e){try{N++;var n="";for(var t of e)if(!k.has(t)){k.add(t);var{loc:i}=t;i&&(n+=i.source.body)}return n}finally{0==--N&&k.clear()}}(n||[]);return{start:0,end:r.length,source:{body:r,name:"GraphQLTada",locationOffset:{line:1,column:1}}}}return i},set loc(e){i=e}}}return e.scalar=function(e,n){return n},e.persisted=function(e,n){return{kind:c,definitions:n?n.definitions:[],documentId:e}},e}function _(e){return 9===e||32===e}V();const F=/[\x00-\x1f\x22\x5c\x7f-\x9f]/g;function w(e){return j[e.charCodeAt(0)]}const j=["\\u0000","\\u0001","\\u0002","\\u0003","\\u0004","\\u0005","\\u0006","\\u0007","\\b","\\t","\\n","\\u000B","\\f","\\r","\\u000E","\\u000F","\\u0010","\\u0011","\\u0012","\\u0013","\\u0014","\\u0015","\\u0016","\\u0017","\\u0018","\\u0019","\\u001A","\\u001B","\\u001C","\\u001D","\\u001E","\\u001F","","",'\\"',"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","\\\\","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","\\u007F","\\u0080","\\u0081","\\u0082","\\u0083","\\u0084","\\u0085","\\u0086","\\u0087","\\u0088","\\u0089","\\u008A","\\u008B","\\u008C","\\u008D","\\u008E","\\u008F","\\u0090","\\u0091","\\u0092","\\u0093","\\u0094","\\u0095","\\u0096","\\u0097","\\u0098","\\u0099","\\u009A","\\u009B","\\u009C","\\u009D","\\u009E","\\u009F"];function P(e,n){if(!Boolean(e))throw new Error(n)}function U(e){return q(e,[])}function q(e,n){switch(typeof e){case"string":return JSON.stringify(e);case"function":return e.name?`[function ${e.name}]`:"[function]";case"object":return function(e,n){if(null===e)return"null";if(n.includes(e))return"[Circular]";const t=[...n,e];if(function(e){return"function"==typeof e.toJSON}(e)){const n=e.toJSON();if(n!==e)return"string"==typeof n?n:q(n,t)}else if(Array.isArray(e))return function(e,n){if(0===e.length)return"[]";if(n.length>2)return"[Array]";const t=Math.min(10,e.length),i=e.length-t,r=[];for(let i=0;i<t;++i)r.push(q(e[i],n));return 1===i?r.push("... 1 more item"):i>1&&r.push(`... ${i} more items`),"["+r.join(", ")+"]"}(e,t);return function(e,n){const t=Object.entries(e);if(0===t.length)return"{}";if(n.length>2)return"["+function(e){const n=Object.prototype.toString.call(e).replace(/^\[object /,"").replace(/]$/,"");if("Object"===n&&"function"==typeof e.constructor){const n=e.constructor.name;if("string"==typeof n&&""!==n)return n}return n}(e)+"]";const i=t.map((([e,t])=>e+": "+q(t,n)));return"{ "+i.join(", ")+" }"}(e,t)}(e,n);default:return String(e)}}Symbol.toStringTag,Symbol.toStringTag;const M={Name:[],Document:["definitions"],OperationDefinition:["name","variableDefinitions","directives","selectionSet"],VariableDefinition:["variable","type","defaultValue","directives"],Variable:["name"],SelectionSet:["selections"],Field:["alias","name","arguments","directives","selectionSet"],Argument:["name","value"],FragmentSpread:["name","directives"],InlineFragment:["typeCondition","directives","selectionSet"],FragmentDefinition:["name","variableDefinitions","typeCondition","directives","selectionSet"],IntValue:[],FloatValue:[],StringValue:[],BooleanValue:[],NullValue:[],EnumValue:[],ListValue:["values"],ObjectValue:["fields"],ObjectField:["name","value"],Directive:["name","arguments"],NamedType:["name"],ListType:["type"],NonNullType:["type"],SchemaDefinition:["description","directives","operationTypes"],OperationTypeDefinition:["type"],ScalarTypeDefinition:["description","name","directives"],ObjectTypeDefinition:["description","name","interfaces","directives","fields"],FieldDefinition:["description","name","arguments","type","directives"],InputValueDefinition:["description","name","type","defaultValue","directives"],InterfaceTypeDefinition:["description","name","interfaces","directives","fields"],UnionTypeDefinition:["description","name","directives","types"],EnumTypeDefinition:["description","name","directives","values"],EnumValueDefinition:["description","name","directives"],InputObjectTypeDefinition:["description","name","directives","fields"],DirectiveDefinition:["description","name","arguments","locations"],SchemaExtension:["directives","operationTypes"],ScalarTypeExtension:["name","directives"],ObjectTypeExtension:["name","interfaces","directives","fields"],InterfaceTypeExtension:["name","interfaces","directives","fields"],UnionTypeExtension:["name","directives","types"],EnumTypeExtension:["name","directives","values"],InputObjectTypeExtension:["name","directives","fields"]},L=new Set(Object.keys(M));function B(e){const n=null==e?void 0:e.kind;return"string"==typeof n&&L.has(n)}var R,K,G;!function(e){e.QUERY="query",e.MUTATION="mutation",e.SUBSCRIPTION="subscription"}(R||(R={})),(G=K||(K={})).NAME="Name",G.DOCUMENT="Document",G.OPERATION_DEFINITION="OperationDefinition",G.VARIABLE_DEFINITION="VariableDefinition",G.SELECTION_SET="SelectionSet",G.FIELD="Field",G.ARGUMENT="Argument",G.FRAGMENT_SPREAD="FragmentSpread",G.INLINE_FRAGMENT="InlineFragment",G.FRAGMENT_DEFINITION="FragmentDefinition",G.VARIABLE="Variable",G.INT="IntValue",G.FLOAT="FloatValue",G.STRING="StringValue",G.BOOLEAN="BooleanValue",G.NULL="NullValue",G.ENUM="EnumValue",G.LIST="ListValue",G.OBJECT="ObjectValue",G.OBJECT_FIELD="ObjectField",G.DIRECTIVE="Directive",G.NAMED_TYPE="NamedType",G.LIST_TYPE="ListType",G.NON_NULL_TYPE="NonNullType",G.SCHEMA_DEFINITION="SchemaDefinition",G.OPERATION_TYPE_DEFINITION="OperationTypeDefinition",G.SCALAR_TYPE_DEFINITION="ScalarTypeDefinition",G.OBJECT_TYPE_DEFINITION="ObjectTypeDefinition",G.FIELD_DEFINITION="FieldDefinition",G.INPUT_VALUE_DEFINITION="InputValueDefinition",G.INTERFACE_TYPE_DEFINITION="InterfaceTypeDefinition",G.UNION_TYPE_DEFINITION="UnionTypeDefinition",G.ENUM_TYPE_DEFINITION="EnumTypeDefinition",G.ENUM_VALUE_DEFINITION="EnumValueDefinition",G.INPUT_OBJECT_TYPE_DEFINITION="InputObjectTypeDefinition",G.DIRECTIVE_DEFINITION="DirectiveDefinition",G.SCHEMA_EXTENSION="SchemaExtension",G.SCALAR_TYPE_EXTENSION="ScalarTypeExtension",G.OBJECT_TYPE_EXTENSION="ObjectTypeExtension",G.INTERFACE_TYPE_EXTENSION="InterfaceTypeExtension",G.UNION_TYPE_EXTENSION="UnionTypeExtension",G.ENUM_TYPE_EXTENSION="EnumTypeExtension",G.INPUT_OBJECT_TYPE_EXTENSION="InputObjectTypeExtension";const Y=Object.freeze({});function J(e,n){const t=e[n];return"object"==typeof t?t:"function"==typeof t?{enter:t,leave:void 0}:{enter:e.enter,leave:e.leave}}const W={Name:{leave:e=>e.value},Variable:{leave:e=>"$"+e.name},Document:{leave:e=>H(e.definitions,"\n\n")},OperationDefinition:{leave(e){const n=X("(",H(e.variableDefinitions,", "),")"),t=H([e.operation,H([e.name,n]),H(e.directives," ")]," ");return("query"===t?"":t+" ")+e.selectionSet}},VariableDefinition:{leave:({variable:e,type:n,defaultValue:t,directives:i})=>e+": "+n+X(" = ",t)+X(" ",H(i," "))},SelectionSet:{leave:({selections:e})=>z(e)},Field:{leave({alias:e,name:n,arguments:t,directives:i,selectionSet:r}){const o=X("",e,": ")+n;let a=o+X("(",H(t,", "),")");return a.length>80&&(a=o+X("(\n",Q(H(t,"\n")),"\n)")),H([a,H(i," "),r]," ")}},Argument:{leave:({name:e,value:n})=>e+": "+n},FragmentSpread:{leave:({name:e,directives:n})=>"..."+e+X(" ",H(n," "))},InlineFragment:{leave:({typeCondition:e,directives:n,selectionSet:t})=>H(["...",X("on ",e),H(n," "),t]," ")},FragmentDefinition:{leave:({name:e,typeCondition:n,variableDefinitions:t,directives:i,selectionSet:r})=>`fragment ${e}${X("(",H(t,", "),")")} on ${n} ${X("",H(i," ")," ")}`+r},IntValue:{leave:({value:e})=>e},FloatValue:{leave:({value:e})=>e},StringValue:{leave:({value:e,block:n})=>n?function(e){const n=e.replace(/"""/g,'\\"""'),t=n.split(/\r\n|[\n\r]/g),i=1===t.length,r=t.length>1&&t.slice(1).every((e=>0===e.length||_(e.charCodeAt(0)))),o=n.endsWith('\\"""'),a=e.endsWith('"')&&!o,s=e.endsWith("\\"),c=a||s,u=!i||e.length>70||c||r||o;let l="";const d=i&&_(e.charCodeAt(0));return(u&&!d||r)&&(l+="\n"),l+=n,(u||c)&&(l+="\n"),'"""'+l+'"""'}(e):`"${e.replace(F,w)}"`},BooleanValue:{leave:({value:e})=>e?"true":"false"},NullValue:{leave:()=>"null"},EnumValue:{leave:({value:e})=>e},ListValue:{leave:({values:e})=>"["+H(e,", ")+"]"},ObjectValue:{leave:({fields:e})=>"{"+H(e,", ")+"}"},ObjectField:{leave:({name:e,value:n})=>e+": "+n},Directive:{leave:({name:e,arguments:n})=>"@"+e+X("(",H(n,", "),")")},NamedType:{leave:({name:e})=>e},ListType:{leave:({type:e})=>"["+e+"]"},NonNullType:{leave:({type:e})=>e+"!"},SchemaDefinition:{leave:({description:e,directives:n,operationTypes:t})=>X("",e,"\n")+H(["schema",H(n," "),z(t)]," ")},OperationTypeDefinition:{leave:({operation:e,type:n})=>e+": "+n},ScalarTypeDefinition:{leave:({description:e,name:n,directives:t})=>X("",e,"\n")+H(["scalar",n,H(t," ")]," ")},ObjectTypeDefinition:{leave:({description:e,name:n,interfaces:t,directives:i,fields:r})=>X("",e,"\n")+H(["type",n,X("implements ",H(t," & ")),H(i," "),z(r)]," ")},FieldDefinition:{leave:({description:e,name:n,arguments:t,type:i,directives:r})=>X("",e,"\n")+n+(Z(t)?X("(\n",Q(H(t,"\n")),"\n)"):X("(",H(t,", "),")"))+": "+i+X(" ",H(r," "))},InputValueDefinition:{leave:({description:e,name:n,type:t,defaultValue:i,directives:r})=>X("",e,"\n")+H([n+": "+t,X("= ",i),H(r," ")]," ")},InterfaceTypeDefinition:{leave:({description:e,name:n,interfaces:t,directives:i,fields:r})=>X("",e,"\n")+H(["interface",n,X("implements ",H(t," & ")),H(i," "),z(r)]," ")},UnionTypeDefinition:{leave:({description:e,name:n,directives:t,types:i})=>X("",e,"\n")+H(["union",n,H(t," "),X("= ",H(i," | "))]," ")},EnumTypeDefinition:{leave:({description:e,name:n,directives:t,values:i})=>X("",e,"\n")+H(["enum",n,H(t," "),z(i)]," ")},EnumValueDefinition:{leave:({description:e,name:n,directives:t})=>X("",e,"\n")+H([n,H(t," ")]," ")},InputObjectTypeDefinition:{leave:({description:e,name:n,directives:t,fields:i})=>X("",e,"\n")+H(["input",n,H(t," "),z(i)]," ")},DirectiveDefinition:{leave:({description:e,name:n,arguments:t,repeatable:i,locations:r})=>X("",e,"\n")+"directive @"+n+(Z(t)?X("(\n",Q(H(t,"\n")),"\n)"):X("(",H(t,", "),")"))+(i?" repeatable":"")+" on "+H(r," | ")},SchemaExtension:{leave:({directives:e,operationTypes:n})=>H(["extend schema",H(e," "),z(n)]," ")},ScalarTypeExtension:{leave:({name:e,directives:n})=>H(["extend scalar",e,H(n," ")]," ")},ObjectTypeExtension:{leave:({name:e,interfaces:n,directives:t,fields:i})=>H(["extend type",e,X("implements ",H(n," & ")),H(t," "),z(i)]," ")},InterfaceTypeExtension:{leave:({name:e,interfaces:n,directives:t,fields:i})=>H(["extend interface",e,X("implements ",H(n," & ")),H(t," "),z(i)]," ")},UnionTypeExtension:{leave:({name:e,directives:n,types:t})=>H(["extend union",e,H(n," "),X("= ",H(t," | "))]," ")},EnumTypeExtension:{leave:({name:e,directives:n,values:t})=>H(["extend enum",e,H(n," "),z(t)]," ")},InputObjectTypeExtension:{leave:({name:e,directives:n,fields:t})=>H(["extend input",e,H(n," "),z(t)]," ")}};function H(e,n=""){var t;return null!==(t=null==e?void 0:e.filter((e=>e)).join(n))&&void 0!==t?t:""}function z(e){return X("{\n",Q(H(e,"\n")),"\n}")}function X(e,n,t=""){return null!=n&&""!==n?e+n+t:""}function Q(e){return X("  ",e.replace(/\n/g,"\n  "))}function Z(e){var n;return null!==(n=null==e?void 0:e.some((e=>e.includes("\n"))))&&void 0!==n&&n}const ee=V();class ne{client;constructor(e,n){const t="live"===e.tradingMode?"4567":"6789",i=e.host.includes(t)?e.host:e.host.replace(/\/$/,`:${t}/`),r={Authorization:`Basic ${e.apiKey} ${e.apiSecret}`};"paper"===e.tradingMode&&(r["x-architect-account-mode"]="paper"),this.client=n({url:`${i}graphql`,headers:r})}parse(e){return ee(e)}async execute(e,n){return new Promise(((t,i)=>{let r;var o;this.client.subscribe({query:(o=e,function(e,n,t=M){const i=new Map;for(const e of Object.values(K))i.set(e,J(n,e));let r,o,a,s=Array.isArray(e),c=[e],u=-1,l=[],d=e;const m=[],y=[];do{u++;const e=u===c.length,h=e&&0!==l.length;if(e){if(o=0===y.length?void 0:m[m.length-1],d=a,a=y.pop(),h)if(s){d=d.slice();let e=0;for(const[n,t]of l){const i=n-e;null===t?(d.splice(i,1),e++):d[i]=t}}else{d=Object.defineProperties({},Object.getOwnPropertyDescriptors(d));for(const[e,n]of l)d[e]=n}u=r.index,c=r.keys,l=r.edits,s=r.inArray,r=r.prev}else if(a){if(o=s?u:c[u],d=a[o],null==d)continue;m.push(o)}let g;if(!Array.isArray(d)){var f,p;B(d)||P(!1,`Invalid AST Node: ${U(d)}.`);const t=e?null===(f=i.get(d.kind))||void 0===f?void 0:f.leave:null===(p=i.get(d.kind))||void 0===p?void 0:p.enter;if(g=null==t?void 0:t.call(n,d,o,a,m,y),g===Y)break;if(!1===g){if(!e){m.pop();continue}}else if(void 0!==g&&(l.push([o,g]),!e)){if(!B(g)){m.pop();continue}d=g}}var v;void 0===g&&h&&l.push([o,d]),e?m.pop():(r={inArray:s,index:u,keys:c,edits:l,prev:r},s=Array.isArray(d),c=s?d:null!==(v=t[d.kind])&&void 0!==v?v:[],u=-1,l=[],a&&y.push(a),a=d)}while(void 0!==r);return 0!==l.length?l[l.length-1][1]:e}(o,W)),variables:n},{next:e=>{r=e.data},error:e=>i(e),complete:()=>t(r)})}))}async debugAsUser(e){return this.execute(ee("mutation DebugAsUser($user: String!) {\n        admin {\n          debugAsUser(user: $user)\n        }\n      }"),{user:e}).then((e=>e.admin.debugAsUser))}async stopDebuggingAsUser(){return this.execute(ee("mutation StopDebuggingAsUser {\n        admin {\n          stopDebuggingAsUser\n        }\n      }")).then((e=>e.admin.stopDebuggingAsUser))}async marketdataVenues(){return this.execute(ee("query MarketdataVenues {\n        config {\n          marketdataVenues\n        }\n      }")).then((e=>e.config.marketdataVenues))}async cmeProductGroupInfo(e,n){return this.execute(ee(`query CmeProductGroupInfo($seriesSymbol: String!) {\n        exchangeSymbology {\n          cmeProductGroupInfo(seriesSymbol: $seriesSymbol) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{seriesSymbol:n}).then((e=>e.exchangeSymbology.cmeProductGroupInfo))}async cmeProductGroupInfos(e,n){return this.execute(ee(`query CmeProductGroupInfos($seriesSymbols: [String!]) {\n        exchangeSymbology {\n          cmeProductGroupInfos(seriesSymbols: $seriesSymbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{seriesSymbols:n}).then((e=>e.exchangeSymbology.cmeProductGroupInfos))}async accountHistory(e,n,t,i,r){return this.execute(ee(`query AccountHistory($account: String!, $fromInclusive: DateTime, $toExclusive: DateTime, $venue: ExecutionVenue) {\n        folio {\n          accountHistory(account: $account, fromInclusive: $fromInclusive, toExclusive: $toExclusive, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,toExclusive:i,venue:r}).then((e=>e.folio.accountHistory))}async accountSummaries(e,n,t){return this.execute(ee(`query AccountSummaries($accounts: [String!], $trader: String) {\n        folio {\n          accountSummaries(accounts: $accounts, trader: $trader) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{accounts:n,trader:t}).then((e=>e.folio.accountSummaries))}async accountSummary(e,n){return this.execute(ee(`query AccountSummary($account: String!) {\n        folio {\n          accountSummary(account: $account) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n}).then((e=>e.folio.accountSummary))}async historicalFills(e,n,t,i,r,o,a,s){return this.execute(ee(`query HistoricalFills($account: String, $fromInclusive: DateTime, $limit: Int, $orderId: OrderId, $toExclusive: DateTime, $trader: String, $venue: ExecutionVenue) {\n        folio {\n          historicalFills(account: $account, fromInclusive: $fromInclusive, limit: $limit, orderId: $orderId, toExclusive: $toExclusive, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,limit:i,orderId:r,toExclusive:o,trader:a,venue:s}).then((e=>e.folio.historicalFills))}async historicalOrders(e,n,t,i,r,o,a,s,c){return this.execute(ee(`query HistoricalOrders($account: String, $fromInclusive: DateTime, $limit: Int, $orderIds: [OrderId!], $parentOrderId: OrderId, $toExclusive: DateTime, $trader: String, $venue: ExecutionVenue) {\n        folio {\n          historicalOrders(account: $account, fromInclusive: $fromInclusive, limit: $limit, orderIds: $orderIds, parentOrderId: $parentOrderId, toExclusive: $toExclusive, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,limit:i,orderIds:r,parentOrderId:o,toExclusive:a,trader:s,venue:c}).then((e=>e.folio.historicalOrders))}async historicalCandles(e,n,t,i,r,o){return this.execute(ee(`query HistoricalCandles($venue: MarketdataVenue!, $symbol: String!, $start: DateTime!, $end: DateTime!, $candleWidth: CandleWidth!) {\n        marketdata {\n          historicalCandles(venue: $venue, symbol: $symbol, start: $start, end: $end, candleWidth: $candleWidth) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{venue:n,symbol:t,start:i,end:r,candleWidth:o}).then((e=>e.marketdata.historicalCandles))}async l2BookSnapshot(e,n,t){return this.execute(ee(`query L2BookSnapshot($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          l2BookSnapshot(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.l2BookSnapshot))}async marketStatus(e,n,t){return this.execute(ee(`query MarketStatus($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          marketStatus(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.marketStatus))}async ticker(e,n,t){return this.execute(ee(`query Ticker($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          ticker(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.ticker))}async tickers(e,n,t,i,r,o){return this.execute(ee(`query Tickers($venue: MarketdataVenue!, $limit: Int, $offset: Int, $sortBy: SortTickersBy, $symbols: [String!]) {\n        marketdata {\n          tickers(venue: $venue, limit: $limit, offset: $offset, sortBy: $sortBy, symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{venue:n,limit:t,offset:i,sortBy:r,symbols:o}).then((e=>e.marketdata.tickers))}async cancelAllOrders(e,n,t){return this.execute(ee("mutation CancelAllOrders($account: String, $executionVenue: String, $trader: String) {\n        oms {\n          cancelAllOrders(account: $account, executionVenue: $executionVenue, trader: $trader)\n        }\n      }"),{account:e,executionVenue:n,trader:t}).then((e=>e.oms.cancelAllOrders))}async cancelOrder(e,n){return this.execute(ee(`mutation CancelOrder($orderId: OrderId!) {\n        oms {\n          cancelOrder(orderId: $orderId) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{orderId:n}).then((e=>e.oms.cancelOrder))}async placeOrder(e,n,t,i,r,o,a,s,c,u,l,d,m,y){return this.execute(ee(`mutation PlaceOrder($timeInForce: TimeInForce!, $symbol: String!, $quantity: Decimal!, $orderType: OrderType!, $dir: Dir!, $account: String, $executionVenue: ExecutionVenue, $goodTilDate: DateTime, $id: OrderId, $limitPrice: Decimal, $postOnly: Boolean, $trader: String, $triggerPrice: Decimal) {\n        oms {\n          placeOrder(timeInForce: $timeInForce, symbol: $symbol, quantity: $quantity, orderType: $orderType, dir: $dir, account: $account, executionVenue: $executionVenue, goodTilDate: $goodTilDate, id: $id, limitPrice: $limitPrice, postOnly: $postOnly, trader: $trader, triggerPrice: $triggerPrice) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{timeInForce:n,symbol:t,quantity:i,orderType:r,dir:o,account:a,executionVenue:s,goodTilDate:c,id:u,limitPrice:l,postOnly:d,trader:m,triggerPrice:y}).then((e=>e.oms.placeOrder))}async openOrders(e,n,t,i,r,o,a){return this.execute(ee(`query OpenOrders($account: String, $orderIds: [OrderId!], $parentOrderId: OrderId, $symbol: String, $trader: String, $venue: ExecutionVenue) {\n        oms {\n          openOrders(account: $account, orderIds: $orderIds, parentOrderId: $parentOrderId, symbol: $symbol, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,orderIds:t,parentOrderId:i,symbol:r,trader:o,venue:a}).then((e=>e.oms.openOrders))}async pendingCancels(e,n,t,i,r,o){return this.execute(ee(`query PendingCancels($account: String, $cancelIds: [Uuid!], $symbol: String, $trader: String, $venue: ExecutionVenue) {\n        oms {\n          pendingCancels(account: $account, cancelIds: $cancelIds, symbol: $symbol, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,cancelIds:t,symbol:i,trader:r,venue:o}).then((e=>e.oms.pendingCancels))}async exchangeSymbology(e){return this.execute(ee(`query ExchangeSymbology {\n        symbology {\n          exchangeSymbology {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.symbology.exchangeSymbology))}async executionInfo(e,n,t){return this.execute(ee(`query ExecutionInfo($symbol: TradableProduct!, $executionVenue: ExecutionVenue!) {\n        symbology {\n          executionInfo(symbol: $symbol, executionVenue: $executionVenue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,executionVenue:t}).then((e=>e.symbology.executionInfo))}async executionInfos(e,n,t){return this.execute(ee(`query ExecutionInfos($executionVenue: ExecutionVenue, $symbols: [TradableProduct!]) {\n        symbology {\n          executionInfos(executionVenue: $executionVenue, symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{executionVenue:n,symbols:t}).then((e=>e.symbology.executionInfos))}async futuresSeries(e){return this.execute(ee("query FuturesSeries($seriesSymbol: String!) {\n        symbology {\n          futuresSeries(seriesSymbol: $seriesSymbol)\n        }\n      }"),{seriesSymbol:e}).then((e=>e.symbology.futuresSeries))}async productInfo(e,n){return this.execute(ee(`query ProductInfo($symbol: String!) {\n        symbology {\n          productInfo(symbol: $symbol) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n}).then((e=>e.symbology.productInfo))}async productInfos(e,n){return this.execute(ee(`query ProductInfos($symbols: [String!]) {\n        symbology {\n          productInfos(symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbols:n}).then((e=>e.symbology.productInfos))}async searchSymbols(e,n,t,i){return this.execute(ee("query SearchSymbols($executionVenue: ExecutionVenue, $limit: Int, $offset: Int, $searchString: String) {\n        symbology {\n          searchSymbols(executionVenue: $executionVenue, limit: $limit, offset: $offset, searchString: $searchString)\n        }\n      }"),{executionVenue:e,limit:n,offset:t,searchString:i}).then((e=>e.symbology.searchSymbols))}async createApiKey(e){return this.execute(ee(`mutation CreateApiKey {\n        user {\n          createApiKey {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.createApiKey))}async createJwt(){return this.execute(ee("mutation CreateJwt {\n        user {\n          createJwt\n        }\n      }")).then((e=>e.user.createJwt))}async enablePaperTrading(){return this.execute(ee("mutation EnablePaperTrading {\n        user {\n          enablePaperTrading\n        }\n      }")).then((e=>e.user.enablePaperTrading))}async removeApiKey(e){return this.execute(ee("mutation RemoveApiKey($apiKey: String!) {\n        user {\n          removeApiKey(apiKey: $apiKey)\n        }\n      }"),{apiKey:e}).then((e=>e.user.removeApiKey))}async account(e,n,t){return this.execute(ee(`query Account($id: Uuid, $name: AccountName) {\n        user {\n          account(id: $id, name: $name) {\n            __typename\n            account { id name }\n            trader\n            permissions { list view trade reduceOrClose setLimits }\n            ${e.join(" ")}\n          }\n        }\n      }`),{id:n,name:t}).then((e=>e.user.account))}async accounts(e){return this.execute(ee(`query Accounts {\n        user {\n          accounts {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.accounts))}async apiKeys(e){return this.execute(ee(`query ApiKeys {\n        user {\n          apiKeys {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.apiKeys))}async canDebugAsUser(){return this.execute(ee("query CanDebugAsUser {\n        user {\n          canDebugAsUser\n        }\n      }")).then((e=>e.user.canDebugAsUser))}async debuggingAsUser(){return this.execute(ee("query DebuggingAsUser {\n        user {\n          debuggingAsUser\n        }\n      }")).then((e=>e.user.debuggingAsUser))}async userEmail(){return this.execute(ee("query UserEmail {\n        user {\n          userEmail\n        }\n      }")).then((e=>e.user.userEmail))}async userId(){return this.execute(ee("query UserId {\n        user {\n          userId\n        }\n      }")).then((e=>e.user.userId))}}var te=t(985);let ie=new Proxy({},{get(e,n){throw new Error("Client is not initialized")},set(e,n,t){throw new Error("Client is not initialized")}});async function re(){const e=await(0,te.$E)("ArchitectApiKey"),n=await(0,te.$E)("ArchitectApiSecret");if(!e)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"api_key has not been input");if(!n)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"api_secret has not been input");var t;return te.$W.apiKey=e,te.$W.apiSecret=n,t=te.$W,ie=new ne(t,r),!0}async function oe(e){throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,"Not implemented")}async function ae(e,n){let t=await ie.ticker(["symbol"],e,n);if(!t)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,"Received bad data from the server, please try again.");try{console.log(t),console.log(t.bidPrice,t.askPrice,t.lastPrice);const e=t.bidPrice?parseFloat(t.bidPrice):NaN;return[[e,t.askPrice?parseFloat(t.askPrice):NaN,t.lastPrice?parseFloat(t.lastPrice):NaN]]}catch(e){throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"Failed to parse bid/ask prices")}}async function se(e,n){let t=await ae(e,n),i=t[0][1],r=t[0][0];return isNaN(r)||isNaN(i)?NaN:(r+i)/2}async function ce(e){return(await ie.searchSymbols(void 0,void 0,void 0,e)).map((e=>[e]))}async function ue(){return(await ie.searchSymbols(void 0,void 0,void 0,"ES 20250321 CME Future"))[0]}async function le(){return(await ie.searchSymbols(void 0,void 0,void 0,"ES 20250321 CME Future")).map((e=>[e]))}return Office.onReady((async e=>{e.host===Office.HostType.Excel&&(await re()?console.log("Client initialized using saved API key/secret"):console.log("Client not initialized because of missing API key or secret"))})),CustomFunctions.associate("INITIALIZECLIENT",re),CustomFunctions.associate("GETMARKETLAST",oe),CustomFunctions.associate("GETMARKETBBO",ae),CustomFunctions.associate("GETMARKETMID",se),CustomFunctions.associate("SEARCHSYMBOLS",ce),CustomFunctions.associate("TESTCLIENT",ue),CustomFunctions.associate("TESTCLIENT2",le),i})()));