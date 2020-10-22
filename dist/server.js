!function(t){var e={};function i(o){if(e[o])return e[o].exports;var n=e[o]={i:o,l:!1,exports:{}};return t[o].call(n.exports,n,n.exports,i),n.l=!0,n.exports}i.m=t,i.c=e,i.d=function(t,e,o){i.o(t,e)||Object.defineProperty(t,e,{enumerable:!0,get:o})},i.r=function(t){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},i.t=function(t,e){if(1&e&&(t=i(t)),8&e)return t;if(4&e&&"object"==typeof t&&t&&t.__esModule)return t;var o=Object.create(null);if(i.r(o),Object.defineProperty(o,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var n in t)i.d(o,n,function(e){return t[e]}.bind(null,n));return o},i.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return i.d(e,"a",e),e},i.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},i.p="",i(i.s=3)}([function(t,e){t.exports=require("express-msteams-host")},function(t,e){t.exports=require("debug")},function(t,e){t.exports=require("botbuilder-dialogs")},function(t,e,i){t.exports=i(4)},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(5),n=i(6),r=i(7),a=i(8),s=i(0),c=i(1)("msteams");c("Initializing Microsoft Teams Express hosted App..."),i(9).config();const l=i(10),u=o(),d=process.env.port||process.env.PORT||3007;u.use(o.json({verify:(t,e,i,o)=>{t.rawBody=i.toString()}})),u.use(o.urlencoded({extended:!0})),u.set("views",r.join(__dirname,"/")),u.use(a("tiny")),u.use("/scripts",o.static(r.join(__dirname,"web/scripts"))),u.use("/assets",o.static(r.join(__dirname,"web/assets"))),u.use(s.MsTeamsApiRouter(l)),u.use(s.MsTeamsPageRouter({root:r.join(__dirname,"web/"),components:l})),u.use("/",o.static(r.join(__dirname,"web/"),{index:"index.html"})),u.set("port",d),n.createServer(u).listen(d,()=>{c("Server running on "+d)})},function(t,e){t.exports=require("express")},function(t,e){t.exports=require("http")},function(t,e){t.exports=require("path")},function(t,e){t.exports=require("morgan")},function(t,e){t.exports=require("dotenv")},function(t,e,i){"use strict";function o(t){for(var i in t)e.hasOwnProperty(i)||(e[i]=t[i])}Object.defineProperty(e,"__esModule",{value:!0}),e.nonce={},o(i(11)),o(i(12))},function(t,e,i){"use strict";var o=this&&this.__decorate||function(t,e,i,o){var n,r=arguments.length,a=r<3?e:null===o?o=Object.getOwnPropertyDescriptor(e,i):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)a=Reflect.decorate(t,e,i,o);else for(var s=t.length-1;s>=0;s--)(n=t[s])&&(a=(r<3?n(a):r>3?n(e,i,a):n(e,i))||a);return r>3&&a&&Object.defineProperty(e,i,a),a};Object.defineProperty(e,"__esModule",{value:!0});const n=i(0);let r=class{};r=o([n.PreventIframe("/acPrototypeTab/index.html")],r),e.AcPrototypeTab=r},function(t,e,i){"use strict";var o=this&&this.__decorate||function(t,e,i,o){var n,r=arguments.length,a=r<3?e:null===o?o=Object.getOwnPropertyDescriptor(e,i):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)a=Reflect.decorate(t,e,i,o);else for(var s=t.length-1;s>=0;s--)(n=t[s])&&(a=(r<3?n(a):r>3?n(e,i,a):n(e,i))||a);return r>3&&a&&Object.defineProperty(e,i,a),a},n=this&&this.__awaiter||function(t,e,i,o){return new(i||(i=Promise))((function(n,r){function a(t){try{c(o.next(t))}catch(t){r(t)}}function s(t){try{c(o.throw(t))}catch(t){r(t)}}function c(t){t.done?n(t.value):new i((function(e){e(t.value)})).then(a,s)}c((o=o.apply(t,e||[])).next())}))};Object.defineProperty(e,"__esModule",{value:!0});const r=i(0),a=i(1),s=i(2),c=i(13),l=i(14),u=i(15),d=i(17),p=i(18),y=i(20),m=i(22);a("msteams");let h=class{constructor(t){this.activityProc=new d.TeamsActivityProcessor,this.conversationState=t,this.dialogState=t.createProperty("dialogState"),this.dialogs=new s.DialogSet(this.dialogState),this.dialogs.add(new l.default("help")),this.activityProc.messageActivityHandler={onMessage:t=>n(this,void 0,void 0,(function*(){const e=d.TeamsContext.from(t);switch(t.activity.type){case c.ActivityTypes.Message:const i=e?e.getActivityTextWithoutMentions().toLowerCase():t.activity.text;if(i.startsWith("hello"))return void(yield t.sendActivity("Oh, hello to you as well!"));if(i.startsWith("help")){const e=yield this.dialogs.createContext(t);yield e.beginDialog("help")}else if(i.includes("card"))try{const e=c.CardFactory.adaptiveCard(u.default);yield t.sendActivity({attachments:[e]})}catch(t){console.log(t)}else console.log(i),yield t.sendActivity("I'm terribly sorry, but my master hasn't trained me to do anything yet...");break;case c.ActivityTypes.Invoke:}return this.conversationState.saveChanges(t)}))},this.activityProc.conversationUpdateActivityHandler={onConversationUpdateActivity:t=>n(this,void 0,void 0,(function*(){if(t.activity.membersAdded&&0!==t.activity.membersAdded.length)for(const e in t.activity.membersAdded)if(t.activity.membersAdded[e].id===t.activity.recipient.id){const e=c.CardFactory.adaptiveCard(u.default);yield t.sendActivity({attachments:[e]})}}))},this.activityProc.messageReactionActivityHandler={onMessageReaction:t=>n(this,void 0,void 0,(function*(){const e=t.activity.reactionsAdded;e&&e[0]&&(yield t.sendActivity({textFormat:"xml",text:`That was an interesting reaction (<b>${e[0].type}</b>)`}))}))},this.activityProc.invokeActivityHandler={onInvoke:t=>n(this,void 0,void 0,(function*(){const e=t,i=c.CardFactory.adaptiveCard(u.default),o=c.CardFactory.adaptiveCard(p.default),n=c.CardFactory.adaptiveCard(y.default),r=c.CardFactory.adaptiveCard(m.default);let a;const s={tab:{type:"continue",value:{cards:[{card:n.content},{card:r.content},{card:o.content},{card:i.content}]}}};switch(e.activity.name){case"task/fetch":a={task:{type:"continue",value:{height:"medium",width:"medium",title:"Quick Actions",card:n}}};break;case"task/submit":a={task:{type:"continue",value:s}};break;case"tab/submit":case"tab/fetch":default:a=s}return{status:200,body:a}}))}}onTurn(t){return n(this,void 0,void 0,(function*(){yield this.activityProc.processIncomingActivity(t)}))}};h=o([r.BotDeclaration("/api/messages",new c.MemoryStorage,process.env.MICROSOFT_APP_ID,process.env.MICROSOFT_APP_PASSWORD),r.PreventIframe("/acPrototypeBot/acProtoBotTab.html")],h),e.AcPrototypeBot=h},function(t,e){t.exports=require("botbuilder")},function(t,e,i){"use strict";var o=this&&this.__awaiter||function(t,e,i,o){return new(i||(i=Promise))((function(n,r){function a(t){try{c(o.next(t))}catch(t){r(t)}}function s(t){try{c(o.throw(t))}catch(t){r(t)}}function c(t){t.done?n(t.value):new i((function(e){e(t.value)})).then(a,s)}c((o=o.apply(t,e||[])).next())}))};Object.defineProperty(e,"__esModule",{value:!0});const n=i(2);class r extends n.Dialog{constructor(t){super(t)}beginDialog(t,e){return o(this,void 0,void 0,(function*(){return t.context.sendActivity("I'm just a friendly but rather stupid bot, and right now I don't have any valuable help for you!"),yield t.endDialog()}))}}e.default=r},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(16);e.default=o},function(t){t.exports={$schema:"http://adaptivecards.io/schemas/adaptive-card.json",type:"AdaptiveCard",version:"1.0",body:[{type:"Image",url:"https://adaptivecards.io/content/cats/1.png",size:"medium"},{type:"TextBlock",spacing:"medium",size:"default",weight:"bolder",text:"Welcome to acPrototype",wrap:!0,maxLines:0},{type:"TextBlock",size:"default",isSubtle:!0,text:"Hello, nice to meet you!",wrap:!0,maxLines:0}],actions:[{type:"Action.OpenUrl",title:"Learn more about Yo Teams",url:"https://aka.ms/yoteams"},{type:"Action.OpenUrl",title:"acPrototype",url:"https://andhillo-relay.servicebus.windows.net/MININT-S5EDEDH"},{type:"Action.Submit",title:"Task Module Invoke",data:{msteams:{type:"task/fetch"}}}]}},function(t,e){t.exports=require("botbuilder-teams")},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(19);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"Container",bleed:!0,style:"emphasis",items:[{type:"TextBlock",text:"Admin Dashboard",color:"Dark",fontType:"Default",size:"Medium",weight:"Bolder"}],spacing:"None"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Configure the Workday App",wrap:!0},{type:"TextBlock",spacing:"None",text:"Decide what features you want enabled/disabled",isSubtle:!0,wrap:!0,fontType:"Default",weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Configure"}]}]}]},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",wrap:!0,text:"View Diagnostics"},{type:"TextBlock",spacing:"None",text:"Troubleshooting a problem or needing to log out",isSubtle:!0,weight:"Lighter",fontType:"Default",wrap:!0}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"View diagnostics"}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Disconnect from Workday",wrap:!0},{type:"TextBlock",spacing:"None",text:"⚠️ Taking this action will sever all connections between Workday and Microsoft Teams for all of your users.",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Disconnect"}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Innovation Services Agreement",wrap:!0},{type:"TextBlock",spacing:"None",text:"Status: ENABLED\nCredentials: Logan McNeil (CHRO, HRExe, Mgr 4000, MatMgr, ProMgr,TalMgr,VPRept)",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Use my credentials"}]}]}],separator:!0,spacing:"Medium"}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(21);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"TextBlock",size:"Medium",weight:"Bolder",text:"Quick Actions"},{type:"ActionSet",actions:[{type:"Action.Submit",title:"Take time off"},{type:"Action.Submit",title:"Look up a coworker"},{type:"Action.Submit",title:"Give feedback"},{type:"Action.Submit",title:"Ask Workday"}],separator:!0}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(23);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"Container",bleed:!0,style:"emphasis",items:[{type:"TextBlock",text:"Manager Dashboard | 3 employees",weight:"Bolder",size:"Medium"}]},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"${creator.profileImage}",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"${creator.name}",wrap:!0},{type:"TextBlock",spacing:"None",text:"🎉   5 year anniversary this week",isSubtle:!0,wrap:!0,fontType:"Default",weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Small"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/36.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",wrap:!0,text:"Alex Edwards"},{type:"TextBlock",spacing:"None",text:"📅  Upcoming time off: Mar 20-25, Apr 1",isSubtle:!0,weight:"Lighter",fontType:"Default"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/40.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Yingdan Huang",wrap:!0},{type:"TextBlock",spacing:"None",text:"🎂  Birthday this week",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}}]);
//# sourceMappingURL=server.js.map