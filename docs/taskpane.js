!function(e,n){"object"==typeof exports&&"object"==typeof module?module.exports=n():"function"==typeof define&&define.amd?define([],n):"object"==typeof exports?exports.taskpane=n():e.taskpane=n()}(this,(()=>(()=>{"use strict";var e={880:(e,n,t)=>{function i(e){const{credentials:n="same-origin",referrer:t,referrerPolicy:i,shouldRetry:o=()=>!1}=e,a=e.fetchFn||fetch,s=e.abortControllerImpl||AbortController,c=(()=>{let e=!1;const n=[];return{get disposed(){return e},onDispose:t=>e?(setTimeout((()=>t()),0),()=>{}):(n.push(t),()=>{n.splice(n.indexOf(t),1)}),dispose(){if(!e){e=!0;for(const e of[...n])e()}}}})();return{subscribe(u,l){if(c.disposed)throw new Error("Client has been disposed");const d=new s,m=c.onDispose((()=>{m(),d.abort()}));return(async()=>{var s;let c=null,m=0;for(;;){if(c){const e=await o(c,m);if(d.signal.aborted)return;if(!e)throw c;m++}try{const o="function"==typeof e.url?await e.url(u):e.url;if(d.signal.aborted)return;const c="function"==typeof e.headers?await e.headers():null!==(s=e.headers)&&void 0!==s?s:{};if(d.signal.aborted)return;let m;try{m=await a(o,{signal:d.signal,method:"POST",headers:Object.assign(Object.assign({},c),{"content-type":"application/json; charset=utf-8",accept:"application/graphql-response+json, application/json"}),credentials:n,referrer:t,referrerPolicy:i,body:JSON.stringify(u)})}catch(e){throw new r(e)}if(!m.ok)throw new r(m);if(!m.body)throw new Error("Missing response body");const y=m.headers.get("content-type");if(!y)throw new Error("Missing response content-type");if(!y.includes("application/graphql-response+json")&&!y.includes("application/json"))throw new Error(`Unsupported response content-type ${y}`);const f=await m.json();return l.next(f),d.abort()}catch(e){if(d.signal.aborted)return;if(!(e instanceof r))throw e;c=e}}})().then((()=>l.complete())).catch((e=>l.error(e))),()=>d.abort()},dispose(){c.dispose()}}}t.d(n,{initializeClient:()=>ie});class r extends Error{constructor(e){let n,t;var i;!function(e){return"object"==typeof e&&null!==e}(i=e)||"boolean"!=typeof i.ok||"number"!=typeof i.status||"string"!=typeof i.statusText?n=e instanceof Error?e.message:String(e):(t=e,n="Server responded with "+e.status+": "+e.statusText),super(n),this.name=this.constructor.name,this.response=t}}var o,a,s="Document",c="FragmentDefinition";class u extends Error{constructor(e,n,t,i,r,o,a){super(e),this.name="GraphQLError",this.message=e,r&&(this.path=r),n&&(this.nodes=Array.isArray(n)?n:[n]),t&&(this.source=t),i&&(this.positions=i),o&&(this.originalError=o);var s=a;if(!s&&o){var c=o.extensions;c&&"object"==typeof c&&(s=c)}this.extensions=s||{}}toJSON(){return{...this,message:this.message}}toString(){return this.message}get[Symbol.toStringTag](){return"GraphQLError"}}function l(e){return new u(`Syntax Error: Unexpected token at ${a} in ${e}`)}function d(e){if(e.lastIndex=a,e.test(o))return o.slice(a,a=e.lastIndex)}var m=/ +(?=[^\s])/y;function y(e){for(var n=e.split("\n"),t="",i=0,r=0,o=n.length-1,a=0;a<n.length;a++)m.lastIndex=0,m.test(n[a])&&(a&&(!i||m.lastIndex<i)&&(i=m.lastIndex),r=r||a,o=a);for(var s=r;s<=o;s++)s!==r&&(t+="\n"),t+=n[s].slice(i).replace(/\\"""/g,'"""');return t}function f(){for(var e=0|o.charCodeAt(a++);9===e||10===e||13===e||32===e||35===e||44===e||65279===e;e=0|o.charCodeAt(a++))if(35===e)for(;10!==(e=o.charCodeAt(a++))&&13!==e;);a--}function p(){for(var e=a,n=0|o.charCodeAt(a++);n>=48&&n<=57||n>=65&&n<=90||95===n||n>=97&&n<=122;n=0|o.charCodeAt(a++));if(e===a-1)throw l("Name");var t=o.slice(e,--a);return f(),t}function v(){return{kind:"Name",value:p()}}var h=/(?:"""|(?:[\s\S]*?[^\\])""")/y,g=/(?:(?:\.\d+)?[eE][+-]?\d+|\.\d+)/y;function b(e){var n;switch(o.charCodeAt(a)){case 91:a++,f();for(var t=[];93!==o.charCodeAt(a);)t.push(b(e));return a++,f(),{kind:"ListValue",values:t};case 123:a++,f();for(var i=[];125!==o.charCodeAt(a);){var r=v();if(58!==o.charCodeAt(a++))throw l("ObjectField");f(),i.push({kind:"ObjectField",name:r,value:b(e)})}return a++,f(),{kind:"ObjectValue",fields:i};case 36:if(e)throw l("Variable");return a++,{kind:"Variable",name:v()};case 34:if(34===o.charCodeAt(a+1)&&34===o.charCodeAt(a+2)){if(a+=3,null==(n=d(h)))throw l("StringValue");return f(),{kind:"StringValue",value:y(n.slice(0,-3)),block:!0}}var s,c=a;a++;var u=!1;for(s=0|o.charCodeAt(a++);92===s&&(a++,u=!0)||10!==s&&13!==s&&34!==s&&s;s=0|o.charCodeAt(a++));if(34!==s)throw l("StringValue");return n=o.slice(c,a),f(),{kind:"StringValue",value:u?JSON.parse(n):n.slice(1,-1),block:!1};case 45:case 48:case 49:case 50:case 51:case 52:case 53:case 54:case 55:case 56:case 57:for(var m,$=a++;(m=0|o.charCodeAt(a++))>=48&&m<=57;);var I=o.slice($,--a);if(46===(m=o.charCodeAt(a))||69===m||101===m){if(null==(n=d(g)))throw l("FloatValue");return f(),{kind:"FloatValue",value:I+n}}return f(),{kind:"IntValue",value:I};case 110:if(117===o.charCodeAt(a+1)&&108===o.charCodeAt(a+2)&&108===o.charCodeAt(a+3))return a+=4,f(),{kind:"NullValue"};break;case 116:if(114===o.charCodeAt(a+1)&&117===o.charCodeAt(a+2)&&101===o.charCodeAt(a+3))return a+=4,f(),{kind:"BooleanValue",value:!0};break;case 102:if(97===o.charCodeAt(a+1)&&108===o.charCodeAt(a+2)&&115===o.charCodeAt(a+3)&&101===o.charCodeAt(a+4))return a+=5,f(),{kind:"BooleanValue",value:!1}}return{kind:"EnumValue",value:p()}}function $(e){if(40===o.charCodeAt(a)){var n=[];a++,f();do{var t=v();if(58!==o.charCodeAt(a++))throw l("Argument");f(),n.push({kind:"Argument",name:t,value:b(e)})}while(41!==o.charCodeAt(a));return a++,f(),n}}function I(e){if(64===o.charCodeAt(a)){var n=[];do{a++,n.push({kind:"Directive",name:v(),arguments:$(e)})}while(64===o.charCodeAt(a));return n}}function E(){for(var e=0;91===o.charCodeAt(a);)e++,a++,f();var n={kind:"NamedType",name:v()};do{if(33===o.charCodeAt(a)&&(a++,f(),n={kind:"NonNullType",type:n}),e){if(93!==o.charCodeAt(a++))throw l("NamedType");f(),n={kind:"ListType",type:n}}}while(e--);return n}function S(){if(123!==o.charCodeAt(a++))throw l("SelectionSet");return f(),T()}function T(){var e=[];do{if(46===o.charCodeAt(a)){if(46!==o.charCodeAt(++a)||46!==o.charCodeAt(++a))throw l("SelectionSet");switch(a++,f(),o.charCodeAt(a)){case 64:e.push({kind:"InlineFragment",typeCondition:void 0,directives:I(!1),selectionSet:S()});break;case 111:110===o.charCodeAt(a+1)?(a+=2,f(),e.push({kind:"InlineFragment",typeCondition:{kind:"NamedType",name:v()},directives:I(!1),selectionSet:S()})):e.push({kind:"FragmentSpread",name:v(),directives:I(!1)});break;case 123:a++,f(),e.push({kind:"InlineFragment",typeCondition:void 0,directives:void 0,selectionSet:T()});break;default:e.push({kind:"FragmentSpread",name:v(),directives:I(!1)})}}else{var n=v(),t=void 0;58===o.charCodeAt(a)&&(a++,f(),t=n,n=v());var i=$(!1),r=I(!1),s=void 0;123===o.charCodeAt(a)&&(a++,f(),s=T()),e.push({kind:"Field",alias:t,name:n,arguments:i,directives:r,selectionSet:s})}}while(125!==o.charCodeAt(a));return a++,f(),{kind:"SelectionSet",selections:e}}function x(){if(f(),40===o.charCodeAt(a)){var e=[];a++,f();do{if(36!==o.charCodeAt(a++))throw l("Variable");var n=v();if(58!==o.charCodeAt(a++))throw l("VariableDefinition");f();var t=E(),i=void 0;61===o.charCodeAt(a)&&(a++,f(),i=b(!0)),f(),e.push({kind:"VariableDefinition",variable:{kind:"Variable",name:n},type:t,defaultValue:i,directives:I(!0)})}while(41!==o.charCodeAt(a));return a++,f(),e}}function A(){var e=v();if(111!==o.charCodeAt(a++)||110!==o.charCodeAt(a++))throw l("FragmentDefinition");return f(),{kind:"FragmentDefinition",name:e,typeCondition:{kind:"NamedType",name:v()},directives:I(!1),selectionSet:S()}}function O(){var e=[];do{if(123===o.charCodeAt(a))a++,f(),e.push({kind:"OperationDefinition",operation:"query",name:void 0,variableDefinitions:void 0,directives:void 0,selectionSet:T()});else{var n=p();switch(n){case"fragment":e.push(A());break;case"query":case"mutation":case"subscription":var t,i=void 0;40!==(t=o.charCodeAt(a))&&64!==t&&123!==t&&(i=v()),e.push({kind:"OperationDefinition",operation:n,name:i,variableDefinitions:x(),directives:I(!1),selectionSet:S()});break;default:throw l("Document")}}}while(a<o.length);return e}function C(e,n){return o=e.body?e.body:e,a=0,f(),n&&n.noLocation?{kind:"Document",definitions:O()}:{kind:"Document",definitions:O(),loc:{start:0,end:o.length,startToken:void 0,endToken:void 0,source:{body:o,name:"graphql.web",locationOffset:{line:1,column:1}}}}}var D=0,N=new Set;function k(){function e(e,n){var t,i,r=C(e).definitions,o=new Set;for(var a of n||[])for(var u of a.definitions)u.kind!==c||o.has(u)||(r.push(u),o.add(u));return(t=r[0].kind===c)&&r[0].directives&&(r[0].directives=r[0].directives.filter((e=>"_unmask"!==e.name.value))),{kind:s,definitions:r,get loc(){if(!i&&t){var r=e+function(e){try{D++;var n="";for(var t of e)if(!N.has(t)){N.add(t);var{loc:i}=t;i&&(n+=i.source.body)}return n}finally{0==--D&&N.clear()}}(n||[]);return{start:0,end:r.length,source:{body:r,name:"GraphQLTada",locationOffset:{line:1,column:1}}}}return i},set loc(e){i=e}}}return e.scalar=function(e,n){return n},e.persisted=function(e,n){return{kind:s,definitions:n?n.definitions:[],documentId:e}},e}function V(e){return 9===e||32===e}k();const _=/[\x00-\x1f\x22\x5c\x7f-\x9f]/g;function F(e){return w[e.charCodeAt(0)]}const w=["\\u0000","\\u0001","\\u0002","\\u0003","\\u0004","\\u0005","\\u0006","\\u0007","\\b","\\t","\\n","\\u000B","\\f","\\r","\\u000E","\\u000F","\\u0010","\\u0011","\\u0012","\\u0013","\\u0014","\\u0015","\\u0016","\\u0017","\\u0018","\\u0019","\\u001A","\\u001B","\\u001C","\\u001D","\\u001E","\\u001F","","",'\\"',"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","\\\\","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","\\u007F","\\u0080","\\u0081","\\u0082","\\u0083","\\u0084","\\u0085","\\u0086","\\u0087","\\u0088","\\u0089","\\u008A","\\u008B","\\u008C","\\u008D","\\u008E","\\u008F","\\u0090","\\u0091","\\u0092","\\u0093","\\u0094","\\u0095","\\u0096","\\u0097","\\u0098","\\u0099","\\u009A","\\u009B","\\u009C","\\u009D","\\u009E","\\u009F"];function j(e,n){if(!Boolean(e))throw new Error(n)}function P(e){return U(e,[])}function U(e,n){switch(typeof e){case"string":return JSON.stringify(e);case"function":return e.name?`[function ${e.name}]`:"[function]";case"object":return function(e,n){if(null===e)return"null";if(n.includes(e))return"[Circular]";const t=[...n,e];if(function(e){return"function"==typeof e.toJSON}(e)){const n=e.toJSON();if(n!==e)return"string"==typeof n?n:U(n,t)}else if(Array.isArray(e))return function(e,n){if(0===e.length)return"[]";if(n.length>2)return"[Array]";const t=Math.min(10,e.length),i=e.length-t,r=[];for(let i=0;i<t;++i)r.push(U(e[i],n));return 1===i?r.push("... 1 more item"):i>1&&r.push(`... ${i} more items`),"["+r.join(", ")+"]"}(e,t);return function(e,n){const t=Object.entries(e);if(0===t.length)return"{}";if(n.length>2)return"["+function(e){const n=Object.prototype.toString.call(e).replace(/^\[object /,"").replace(/]$/,"");if("Object"===n&&"function"==typeof e.constructor){const n=e.constructor.name;if("string"==typeof n&&""!==n)return n}return n}(e)+"]";const i=t.map((([e,t])=>e+": "+U(t,n)));return"{ "+i.join(", ")+" }"}(e,t)}(e,n);default:return String(e)}}Symbol.toStringTag,Symbol.toStringTag;const q={Name:[],Document:["definitions"],OperationDefinition:["name","variableDefinitions","directives","selectionSet"],VariableDefinition:["variable","type","defaultValue","directives"],Variable:["name"],SelectionSet:["selections"],Field:["alias","name","arguments","directives","selectionSet"],Argument:["name","value"],FragmentSpread:["name","directives"],InlineFragment:["typeCondition","directives","selectionSet"],FragmentDefinition:["name","variableDefinitions","typeCondition","directives","selectionSet"],IntValue:[],FloatValue:[],StringValue:[],BooleanValue:[],NullValue:[],EnumValue:[],ListValue:["values"],ObjectValue:["fields"],ObjectField:["name","value"],Directive:["name","arguments"],NamedType:["name"],ListType:["type"],NonNullType:["type"],SchemaDefinition:["description","directives","operationTypes"],OperationTypeDefinition:["type"],ScalarTypeDefinition:["description","name","directives"],ObjectTypeDefinition:["description","name","interfaces","directives","fields"],FieldDefinition:["description","name","arguments","type","directives"],InputValueDefinition:["description","name","type","defaultValue","directives"],InterfaceTypeDefinition:["description","name","interfaces","directives","fields"],UnionTypeDefinition:["description","name","directives","types"],EnumTypeDefinition:["description","name","directives","values"],EnumValueDefinition:["description","name","directives"],InputObjectTypeDefinition:["description","name","directives","fields"],DirectiveDefinition:["description","name","arguments","locations"],SchemaExtension:["directives","operationTypes"],ScalarTypeExtension:["name","directives"],ObjectTypeExtension:["name","interfaces","directives","fields"],InterfaceTypeExtension:["name","interfaces","directives","fields"],UnionTypeExtension:["name","directives","types"],EnumTypeExtension:["name","directives","values"],InputObjectTypeExtension:["name","directives","fields"]},L=new Set(Object.keys(q));function M(e){const n=null==e?void 0:e.kind;return"string"==typeof n&&L.has(n)}var B,R,K;!function(e){e.QUERY="query",e.MUTATION="mutation",e.SUBSCRIPTION="subscription"}(B||(B={})),(K=R||(R={})).NAME="Name",K.DOCUMENT="Document",K.OPERATION_DEFINITION="OperationDefinition",K.VARIABLE_DEFINITION="VariableDefinition",K.SELECTION_SET="SelectionSet",K.FIELD="Field",K.ARGUMENT="Argument",K.FRAGMENT_SPREAD="FragmentSpread",K.INLINE_FRAGMENT="InlineFragment",K.FRAGMENT_DEFINITION="FragmentDefinition",K.VARIABLE="Variable",K.INT="IntValue",K.FLOAT="FloatValue",K.STRING="StringValue",K.BOOLEAN="BooleanValue",K.NULL="NullValue",K.ENUM="EnumValue",K.LIST="ListValue",K.OBJECT="ObjectValue",K.OBJECT_FIELD="ObjectField",K.DIRECTIVE="Directive",K.NAMED_TYPE="NamedType",K.LIST_TYPE="ListType",K.NON_NULL_TYPE="NonNullType",K.SCHEMA_DEFINITION="SchemaDefinition",K.OPERATION_TYPE_DEFINITION="OperationTypeDefinition",K.SCALAR_TYPE_DEFINITION="ScalarTypeDefinition",K.OBJECT_TYPE_DEFINITION="ObjectTypeDefinition",K.FIELD_DEFINITION="FieldDefinition",K.INPUT_VALUE_DEFINITION="InputValueDefinition",K.INTERFACE_TYPE_DEFINITION="InterfaceTypeDefinition",K.UNION_TYPE_DEFINITION="UnionTypeDefinition",K.ENUM_TYPE_DEFINITION="EnumTypeDefinition",K.ENUM_VALUE_DEFINITION="EnumValueDefinition",K.INPUT_OBJECT_TYPE_DEFINITION="InputObjectTypeDefinition",K.DIRECTIVE_DEFINITION="DirectiveDefinition",K.SCHEMA_EXTENSION="SchemaExtension",K.SCALAR_TYPE_EXTENSION="ScalarTypeExtension",K.OBJECT_TYPE_EXTENSION="ObjectTypeExtension",K.INTERFACE_TYPE_EXTENSION="InterfaceTypeExtension",K.UNION_TYPE_EXTENSION="UnionTypeExtension",K.ENUM_TYPE_EXTENSION="EnumTypeExtension",K.INPUT_OBJECT_TYPE_EXTENSION="InputObjectTypeExtension";const G=Object.freeze({});function Y(e,n){const t=e[n];return"object"==typeof t?t:"function"==typeof t?{enter:t,leave:void 0}:{enter:e.enter,leave:e.leave}}const J={Name:{leave:e=>e.value},Variable:{leave:e=>"$"+e.name},Document:{leave:e=>H(e.definitions,"\n\n")},OperationDefinition:{leave(e){const n=z("(",H(e.variableDefinitions,", "),")"),t=H([e.operation,H([e.name,n]),H(e.directives," ")]," ");return("query"===t?"":t+" ")+e.selectionSet}},VariableDefinition:{leave:({variable:e,type:n,defaultValue:t,directives:i})=>e+": "+n+z(" = ",t)+z(" ",H(i," "))},SelectionSet:{leave:({selections:e})=>W(e)},Field:{leave({alias:e,name:n,arguments:t,directives:i,selectionSet:r}){const o=z("",e,": ")+n;let a=o+z("(",H(t,", "),")");return a.length>80&&(a=o+z("(\n",X(H(t,"\n")),"\n)")),H([a,H(i," "),r]," ")}},Argument:{leave:({name:e,value:n})=>e+": "+n},FragmentSpread:{leave:({name:e,directives:n})=>"..."+e+z(" ",H(n," "))},InlineFragment:{leave:({typeCondition:e,directives:n,selectionSet:t})=>H(["...",z("on ",e),H(n," "),t]," ")},FragmentDefinition:{leave:({name:e,typeCondition:n,variableDefinitions:t,directives:i,selectionSet:r})=>`fragment ${e}${z("(",H(t,", "),")")} on ${n} ${z("",H(i," ")," ")}`+r},IntValue:{leave:({value:e})=>e},FloatValue:{leave:({value:e})=>e},StringValue:{leave:({value:e,block:n})=>n?function(e){const n=e.replace(/"""/g,'\\"""'),t=n.split(/\r\n|[\n\r]/g),i=1===t.length,r=t.length>1&&t.slice(1).every((e=>0===e.length||V(e.charCodeAt(0)))),o=n.endsWith('\\"""'),a=e.endsWith('"')&&!o,s=e.endsWith("\\"),c=a||s,u=!i||e.length>70||c||r||o;let l="";const d=i&&V(e.charCodeAt(0));return(u&&!d||r)&&(l+="\n"),l+=n,(u||c)&&(l+="\n"),'"""'+l+'"""'}(e):`"${e.replace(_,F)}"`},BooleanValue:{leave:({value:e})=>e?"true":"false"},NullValue:{leave:()=>"null"},EnumValue:{leave:({value:e})=>e},ListValue:{leave:({values:e})=>"["+H(e,", ")+"]"},ObjectValue:{leave:({fields:e})=>"{"+H(e,", ")+"}"},ObjectField:{leave:({name:e,value:n})=>e+": "+n},Directive:{leave:({name:e,arguments:n})=>"@"+e+z("(",H(n,", "),")")},NamedType:{leave:({name:e})=>e},ListType:{leave:({type:e})=>"["+e+"]"},NonNullType:{leave:({type:e})=>e+"!"},SchemaDefinition:{leave:({description:e,directives:n,operationTypes:t})=>z("",e,"\n")+H(["schema",H(n," "),W(t)]," ")},OperationTypeDefinition:{leave:({operation:e,type:n})=>e+": "+n},ScalarTypeDefinition:{leave:({description:e,name:n,directives:t})=>z("",e,"\n")+H(["scalar",n,H(t," ")]," ")},ObjectTypeDefinition:{leave:({description:e,name:n,interfaces:t,directives:i,fields:r})=>z("",e,"\n")+H(["type",n,z("implements ",H(t," & ")),H(i," "),W(r)]," ")},FieldDefinition:{leave:({description:e,name:n,arguments:t,type:i,directives:r})=>z("",e,"\n")+n+(Q(t)?z("(\n",X(H(t,"\n")),"\n)"):z("(",H(t,", "),")"))+": "+i+z(" ",H(r," "))},InputValueDefinition:{leave:({description:e,name:n,type:t,defaultValue:i,directives:r})=>z("",e,"\n")+H([n+": "+t,z("= ",i),H(r," ")]," ")},InterfaceTypeDefinition:{leave:({description:e,name:n,interfaces:t,directives:i,fields:r})=>z("",e,"\n")+H(["interface",n,z("implements ",H(t," & ")),H(i," "),W(r)]," ")},UnionTypeDefinition:{leave:({description:e,name:n,directives:t,types:i})=>z("",e,"\n")+H(["union",n,H(t," "),z("= ",H(i," | "))]," ")},EnumTypeDefinition:{leave:({description:e,name:n,directives:t,values:i})=>z("",e,"\n")+H(["enum",n,H(t," "),W(i)]," ")},EnumValueDefinition:{leave:({description:e,name:n,directives:t})=>z("",e,"\n")+H([n,H(t," ")]," ")},InputObjectTypeDefinition:{leave:({description:e,name:n,directives:t,fields:i})=>z("",e,"\n")+H(["input",n,H(t," "),W(i)]," ")},DirectiveDefinition:{leave:({description:e,name:n,arguments:t,repeatable:i,locations:r})=>z("",e,"\n")+"directive @"+n+(Q(t)?z("(\n",X(H(t,"\n")),"\n)"):z("(",H(t,", "),")"))+(i?" repeatable":"")+" on "+H(r," | ")},SchemaExtension:{leave:({directives:e,operationTypes:n})=>H(["extend schema",H(e," "),W(n)]," ")},ScalarTypeExtension:{leave:({name:e,directives:n})=>H(["extend scalar",e,H(n," ")]," ")},ObjectTypeExtension:{leave:({name:e,interfaces:n,directives:t,fields:i})=>H(["extend type",e,z("implements ",H(n," & ")),H(t," "),W(i)]," ")},InterfaceTypeExtension:{leave:({name:e,interfaces:n,directives:t,fields:i})=>H(["extend interface",e,z("implements ",H(n," & ")),H(t," "),W(i)]," ")},UnionTypeExtension:{leave:({name:e,directives:n,types:t})=>H(["extend union",e,H(n," "),z("= ",H(t," | "))]," ")},EnumTypeExtension:{leave:({name:e,directives:n,values:t})=>H(["extend enum",e,H(n," "),W(t)]," ")},InputObjectTypeExtension:{leave:({name:e,directives:n,fields:t})=>H(["extend input",e,H(n," "),W(t)]," ")}};function H(e,n=""){var t;return null!==(t=null==e?void 0:e.filter((e=>e)).join(n))&&void 0!==t?t:""}function W(e){return z("{\n",X(H(e,"\n")),"\n}")}function z(e,n,t=""){return null!=n&&""!==n?e+n+t:""}function X(e){return z("  ",e.replace(/\n/g,"\n  "))}function Q(e){var n;return null!==(n=null==e?void 0:e.some((e=>e.includes("\n"))))&&void 0!==n&&n}const Z=k();class ee{client;constructor(e,n){const t="live"===e.tradingMode?"4567":"6789",i=e.host.includes(t)?e.host:e.host.replace(/\/$/,`:${t}/`),r={Authorization:`Basic ${e.apiKey} ${e.apiSecret}`};"paper"===e.tradingMode&&(r["x-architect-account-mode"]="paper"),this.client=n({url:`${i}graphql`,headers:r})}parse(e){return Z(e)}async execute(e,n){return new Promise(((t,i)=>{let r;var o;this.client.subscribe({query:(o=e,function(e,n,t=q){const i=new Map;for(const e of Object.values(R))i.set(e,Y(n,e));let r,o,a,s=Array.isArray(e),c=[e],u=-1,l=[],d=e;const m=[],y=[];do{u++;const e=u===c.length,h=e&&0!==l.length;if(e){if(o=0===y.length?void 0:m[m.length-1],d=a,a=y.pop(),h)if(s){d=d.slice();let e=0;for(const[n,t]of l){const i=n-e;null===t?(d.splice(i,1),e++):d[i]=t}}else{d=Object.defineProperties({},Object.getOwnPropertyDescriptors(d));for(const[e,n]of l)d[e]=n}u=r.index,c=r.keys,l=r.edits,s=r.inArray,r=r.prev}else if(a){if(o=s?u:c[u],d=a[o],null==d)continue;m.push(o)}let g;if(!Array.isArray(d)){var f,p;M(d)||j(!1,`Invalid AST Node: ${P(d)}.`);const t=e?null===(f=i.get(d.kind))||void 0===f?void 0:f.leave:null===(p=i.get(d.kind))||void 0===p?void 0:p.enter;if(g=null==t?void 0:t.call(n,d,o,a,m,y),g===G)break;if(!1===g){if(!e){m.pop();continue}}else if(void 0!==g&&(l.push([o,g]),!e)){if(!M(g)){m.pop();continue}d=g}}var v;void 0===g&&h&&l.push([o,d]),e?m.pop():(r={inArray:s,index:u,keys:c,edits:l,prev:r},s=Array.isArray(d),c=s?d:null!==(v=t[d.kind])&&void 0!==v?v:[],u=-1,l=[],a&&y.push(a),a=d)}while(void 0!==r);return 0!==l.length?l[l.length-1][1]:e}(o,J)),variables:n},{next:e=>{r=e.data},error:e=>i(e),complete:()=>t(r)})}))}async debugAsUser(e){return this.execute(Z("mutation DebugAsUser($user: String!) {\n        admin {\n          debugAsUser(user: $user)\n        }\n      }"),{user:e}).then((e=>e.admin.debugAsUser))}async stopDebuggingAsUser(){return this.execute(Z("mutation StopDebuggingAsUser {\n        admin {\n          stopDebuggingAsUser\n        }\n      }")).then((e=>e.admin.stopDebuggingAsUser))}async marketdataVenues(){return this.execute(Z("query MarketdataVenues {\n        config {\n          marketdataVenues\n        }\n      }")).then((e=>e.config.marketdataVenues))}async cmeProductGroupInfo(e,n){return this.execute(Z(`query CmeProductGroupInfo($seriesSymbol: String!) {\n        exchangeSymbology {\n          cmeProductGroupInfo(seriesSymbol: $seriesSymbol) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{seriesSymbol:n}).then((e=>e.exchangeSymbology.cmeProductGroupInfo))}async cmeProductGroupInfos(e,n){return this.execute(Z(`query CmeProductGroupInfos($seriesSymbols: [String!]) {\n        exchangeSymbology {\n          cmeProductGroupInfos(seriesSymbols: $seriesSymbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{seriesSymbols:n}).then((e=>e.exchangeSymbology.cmeProductGroupInfos))}async accountHistory(e,n,t,i,r){return this.execute(Z(`query AccountHistory($account: String!, $fromInclusive: DateTime, $toExclusive: DateTime, $venue: ExecutionVenue) {\n        folio {\n          accountHistory(account: $account, fromInclusive: $fromInclusive, toExclusive: $toExclusive, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,toExclusive:i,venue:r}).then((e=>e.folio.accountHistory))}async accountSummaries(e,n,t){return this.execute(Z(`query AccountSummaries($accounts: [String!], $trader: String) {\n        folio {\n          accountSummaries(accounts: $accounts, trader: $trader) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{accounts:n,trader:t}).then((e=>e.folio.accountSummaries))}async accountSummary(e,n){return this.execute(Z(`query AccountSummary($account: String!) {\n        folio {\n          accountSummary(account: $account) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n}).then((e=>e.folio.accountSummary))}async historicalFills(e,n,t,i,r,o,a,s){return this.execute(Z(`query HistoricalFills($account: String, $fromInclusive: DateTime, $limit: Int, $orderId: OrderId, $toExclusive: DateTime, $trader: String, $venue: ExecutionVenue) {\n        folio {\n          historicalFills(account: $account, fromInclusive: $fromInclusive, limit: $limit, orderId: $orderId, toExclusive: $toExclusive, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,limit:i,orderId:r,toExclusive:o,trader:a,venue:s}).then((e=>e.folio.historicalFills))}async historicalOrders(e,n,t,i,r,o,a,s,c){return this.execute(Z(`query HistoricalOrders($account: String, $fromInclusive: DateTime, $limit: Int, $orderIds: [OrderId!], $parentOrderId: OrderId, $toExclusive: DateTime, $trader: String, $venue: ExecutionVenue) {\n        folio {\n          historicalOrders(account: $account, fromInclusive: $fromInclusive, limit: $limit, orderIds: $orderIds, parentOrderId: $parentOrderId, toExclusive: $toExclusive, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,fromInclusive:t,limit:i,orderIds:r,parentOrderId:o,toExclusive:a,trader:s,venue:c}).then((e=>e.folio.historicalOrders))}async historicalCandles(e,n,t,i,r,o){return this.execute(Z(`query HistoricalCandles($venue: MarketdataVenue!, $symbol: String!, $start: DateTime!, $end: DateTime!, $candleWidth: CandleWidth!) {\n        marketdata {\n          historicalCandles(venue: $venue, symbol: $symbol, start: $start, end: $end, candleWidth: $candleWidth) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{venue:n,symbol:t,start:i,end:r,candleWidth:o}).then((e=>e.marketdata.historicalCandles))}async l2BookSnapshot(e,n,t){return this.execute(Z(`query L2BookSnapshot($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          l2BookSnapshot(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.l2BookSnapshot))}async marketStatus(e,n,t){return this.execute(Z(`query MarketStatus($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          marketStatus(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.marketStatus))}async ticker(e,n,t){return this.execute(Z(`query Ticker($symbol: String!, $venue: MarketdataVenue) {\n        marketdata {\n          ticker(symbol: $symbol, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,venue:t}).then((e=>e.marketdata.ticker))}async tickers(e,n,t,i,r,o){return this.execute(Z(`query Tickers($venue: MarketdataVenue!, $limit: Int, $offset: Int, $sortBy: SortTickersBy, $symbols: [String!]) {\n        marketdata {\n          tickers(venue: $venue, limit: $limit, offset: $offset, sortBy: $sortBy, symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{venue:n,limit:t,offset:i,sortBy:r,symbols:o}).then((e=>e.marketdata.tickers))}async cancelAllOrders(e,n,t){return this.execute(Z("mutation CancelAllOrders($account: String, $executionVenue: String, $trader: String) {\n        oms {\n          cancelAllOrders(account: $account, executionVenue: $executionVenue, trader: $trader)\n        }\n      }"),{account:e,executionVenue:n,trader:t}).then((e=>e.oms.cancelAllOrders))}async cancelOrder(e,n){return this.execute(Z(`mutation CancelOrder($orderId: OrderId!) {\n        oms {\n          cancelOrder(orderId: $orderId) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{orderId:n}).then((e=>e.oms.cancelOrder))}async placeOrder(e,n,t,i,r,o,a,s,c,u,l,d,m,y){return this.execute(Z(`mutation PlaceOrder($timeInForce: TimeInForce!, $symbol: String!, $quantity: Decimal!, $orderType: OrderType!, $dir: Dir!, $account: String, $executionVenue: ExecutionVenue, $goodTilDate: DateTime, $id: OrderId, $limitPrice: Decimal, $postOnly: Boolean, $trader: String, $triggerPrice: Decimal) {\n        oms {\n          placeOrder(timeInForce: $timeInForce, symbol: $symbol, quantity: $quantity, orderType: $orderType, dir: $dir, account: $account, executionVenue: $executionVenue, goodTilDate: $goodTilDate, id: $id, limitPrice: $limitPrice, postOnly: $postOnly, trader: $trader, triggerPrice: $triggerPrice) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{timeInForce:n,symbol:t,quantity:i,orderType:r,dir:o,account:a,executionVenue:s,goodTilDate:c,id:u,limitPrice:l,postOnly:d,trader:m,triggerPrice:y}).then((e=>e.oms.placeOrder))}async openOrders(e,n,t,i,r,o,a){return this.execute(Z(`query OpenOrders($account: String, $orderIds: [OrderId!], $parentOrderId: OrderId, $symbol: String, $trader: String, $venue: ExecutionVenue) {\n        oms {\n          openOrders(account: $account, orderIds: $orderIds, parentOrderId: $parentOrderId, symbol: $symbol, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,orderIds:t,parentOrderId:i,symbol:r,trader:o,venue:a}).then((e=>e.oms.openOrders))}async pendingCancels(e,n,t,i,r,o){return this.execute(Z(`query PendingCancels($account: String, $cancelIds: [Uuid!], $symbol: String, $trader: String, $venue: ExecutionVenue) {\n        oms {\n          pendingCancels(account: $account, cancelIds: $cancelIds, symbol: $symbol, trader: $trader, venue: $venue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{account:n,cancelIds:t,symbol:i,trader:r,venue:o}).then((e=>e.oms.pendingCancels))}async exchangeSymbology(e){return this.execute(Z(`query ExchangeSymbology {\n        symbology {\n          exchangeSymbology {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.symbology.exchangeSymbology))}async executionInfo(e,n,t){return this.execute(Z(`query ExecutionInfo($symbol: TradableProduct!, $executionVenue: ExecutionVenue!) {\n        symbology {\n          executionInfo(symbol: $symbol, executionVenue: $executionVenue) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n,executionVenue:t}).then((e=>e.symbology.executionInfo))}async executionInfos(e,n,t){return this.execute(Z(`query ExecutionInfos($executionVenue: ExecutionVenue, $symbols: [TradableProduct!]) {\n        symbology {\n          executionInfos(executionVenue: $executionVenue, symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{executionVenue:n,symbols:t}).then((e=>e.symbology.executionInfos))}async futuresSeries(e){return this.execute(Z("query FuturesSeries($seriesSymbol: String!) {\n        symbology {\n          futuresSeries(seriesSymbol: $seriesSymbol)\n        }\n      }"),{seriesSymbol:e}).then((e=>e.symbology.futuresSeries))}async productInfo(e,n){return this.execute(Z(`query ProductInfo($symbol: String!) {\n        symbology {\n          productInfo(symbol: $symbol) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbol:n}).then((e=>e.symbology.productInfo))}async productInfos(e,n){return this.execute(Z(`query ProductInfos($symbols: [String!]) {\n        symbology {\n          productInfos(symbols: $symbols) {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`),{symbols:n}).then((e=>e.symbology.productInfos))}async searchSymbols(e,n,t,i){return this.execute(Z("query SearchSymbols($executionVenue: ExecutionVenue, $limit: Int, $offset: Int, $searchString: String) {\n        symbology {\n          searchSymbols(executionVenue: $executionVenue, limit: $limit, offset: $offset, searchString: $searchString)\n        }\n      }"),{executionVenue:e,limit:n,offset:t,searchString:i}).then((e=>e.symbology.searchSymbols))}async createApiKey(e){return this.execute(Z(`mutation CreateApiKey {\n        user {\n          createApiKey {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.createApiKey))}async createJwt(){return this.execute(Z("mutation CreateJwt {\n        user {\n          createJwt\n        }\n      }")).then((e=>e.user.createJwt))}async enablePaperTrading(){return this.execute(Z("mutation EnablePaperTrading {\n        user {\n          enablePaperTrading\n        }\n      }")).then((e=>e.user.enablePaperTrading))}async removeApiKey(e){return this.execute(Z("mutation RemoveApiKey($apiKey: String!) {\n        user {\n          removeApiKey(apiKey: $apiKey)\n        }\n      }"),{apiKey:e}).then((e=>e.user.removeApiKey))}async account(e,n,t){return this.execute(Z(`query Account($id: Uuid, $name: AccountName) {\n        user {\n          account(id: $id, name: $name) {\n            __typename\n            account { id name }\n            trader\n            permissions { list view trade reduceOrClose setLimits }\n            ${e.join(" ")}\n          }\n        }\n      }`),{id:n,name:t}).then((e=>e.user.account))}async accounts(e){return this.execute(Z(`query Accounts {\n        user {\n          accounts {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.accounts))}async apiKeys(e){return this.execute(Z(`query ApiKeys {\n        user {\n          apiKeys {\n            __typename ${e.join(" ")}\n          }\n        }\n      }`)).then((e=>e.user.apiKeys))}async canDebugAsUser(){return this.execute(Z("query CanDebugAsUser {\n        user {\n          canDebugAsUser\n        }\n      }")).then((e=>e.user.canDebugAsUser))}async debuggingAsUser(){return this.execute(Z("query DebuggingAsUser {\n        user {\n          debuggingAsUser\n        }\n      }")).then((e=>e.user.debuggingAsUser))}async userEmail(){return this.execute(Z("query UserEmail {\n        user {\n          userEmail\n        }\n      }")).then((e=>e.user.userEmail))}async userId(){return this.execute(Z("query UserId {\n        user {\n          userId\n        }\n      }")).then((e=>e.user.userId))}}var ne=t(985);let te=new Proxy({},{get(e,n){throw new Error("Client is not initialized")},set(e,n,t){throw new Error("Client is not initialized")}});async function ie(){const e=await(0,ne.$E)("ArchitectApiKey"),n=await(0,ne.$E)("ArchitectApiSecret");if(!e)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"api_key has not been input");if(!n)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"api_secret has not been input");var t;return ne.$W.apiKey=e,ne.$W.apiSecret=n,t=ne.$W,te=new ee(t,i),!0}async function re(e,n){let t=await te.ticker(["symbol"],e,n);if(!t)throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,"Received bad data from the server, please try again.");try{console.log(t),console.log(t.bidPrice,t.askPrice,t.lastPrice);const e=t.bidPrice?parseFloat(t.bidPrice):NaN;return[[e,t.askPrice?parseFloat(t.askPrice):NaN,t.lastPrice?parseFloat(t.lastPrice):NaN]]}catch(e){throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,"Failed to parse bid/ask prices")}}Office.onReady((async e=>{e.host===Office.HostType.Excel&&(await ie()?console.log("Client initialized using saved API key/secret"):console.log("Client not initialized because of missing API key or secret"))})),CustomFunctions.associate("INITIALIZECLIENT",ie),CustomFunctions.associate("GETMARKETLAST",(async function(e){throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,"Not implemented")})),CustomFunctions.associate("GETMARKETBBO",re),CustomFunctions.associate("GETMARKETMID",(async function(e,n){let t=await re(e,n),i=t[0][1],r=t[0][0];return isNaN(r)||isNaN(i)?NaN:(r+i)/2})),CustomFunctions.associate("SEARCHSYMBOLS",(async function(e){return(await te.searchSymbols(void 0,void 0,void 0,e)).map((e=>[e]))})),CustomFunctions.associate("TESTCLIENT",(async function(){return(await te.searchSymbols(void 0,void 0,void 0,"ES 20250321 CME Future"))[0]})),CustomFunctions.associate("TESTCLIENT2",(async function(){return(await te.searchSymbols(void 0,void 0,void 0,"ES 20250321 CME Future")).map((e=>[e]))}))},985:(e,n,t)=>{t.d(n,{$E:()=>o,$W:()=>i,ni:()=>r});let i={host:"https://app.architect.co/",apiKey:"",apiSecret:"",tradingMode:"live"};async function r(e,n){if("undefined"!=typeof Office&&Office.context&&"undefined"!=typeof OfficeRuntime)await OfficeRuntime.storage.setItem(e,n);else{if("undefined"==typeof localStorage)throw new Error("No available storage method to set to.");localStorage.setItem(e,n)}}async function o(e){if("undefined"!=typeof Office&&Office.context&&"undefined"!=typeof OfficeRuntime)return await OfficeRuntime.storage.getItem(e);if("undefined"!=typeof localStorage)return localStorage.getItem(e);throw new Error("No available storage method to get from.")}}},n={};function t(i){var r=n[i];if(void 0!==r)return r.exports;var o=n[i]={exports:{}};return e[i](o,o.exports,t),o.exports}t.d=(e,n)=>{for(var i in n)t.o(n,i)&&!t.o(e,i)&&Object.defineProperty(e,i,{enumerable:!0,get:n[i]})},t.o=(e,n)=>Object.prototype.hasOwnProperty.call(e,n),t.r=e=>{"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})};var i={};t.r(i);var r=t(880),o=t(985);return Office.onReady((async()=>{const e=document.getElementById("api-form");function n(e){return(e||"").trim()}e.addEventListener("submit",(async t=>{t.preventDefault();const i=new FormData(e),a=n(i.get("apiKey")),s=n(i.get("apiSecret")),c=document.getElementById("status");if(a&&s)try{(0,o.ni)("ArchitectApiKey",a),(0,o.ni)("ArchitectApiSecret",s);let e=await(0,r.initializeClient)();c.textContent=e?"Credentials saved! Client initialized!":"Credentials saved! However, Client was NOT successfully initialized!"}catch(e){c.textContent=`Error: ${e.message}`}else c.textContent="API Key and Secret are required."}))})),i})()));