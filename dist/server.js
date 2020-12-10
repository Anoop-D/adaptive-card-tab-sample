!function(t){var e={};function i(o){if(e[o])return e[o].exports;var n=e[o]={i:o,l:!1,exports:{}};return t[o].call(n.exports,n,n.exports,i),n.l=!0,n.exports}i.m=t,i.c=e,i.d=function(t,e,o){i.o(t,e)||Object.defineProperty(t,e,{enumerable:!0,get:o})},i.r=function(t){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},i.t=function(t,e){if(1&e&&(t=i(t)),8&e)return t;if(4&e&&"object"==typeof t&&t&&t.__esModule)return t;var o=Object.create(null);if(i.r(o),Object.defineProperty(o,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var n in t)i.d(o,n,function(e){return t[e]}.bind(null,n));return o},i.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return i.d(e,"a",e),e},i.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},i.p="",i(i.s=3)}([function(t,e){t.exports=require("express-msteams-host")},function(t,e){t.exports=require("debug")},function(t,e){t.exports=require("botbuilder-dialogs")},function(t,e,i){t.exports=i(4)},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(5),n=i(6),a=i(7),s=i(8),r=i(0),c=i(1)("msteams");c("Initializing Microsoft Teams Express hosted App..."),i(9).config();const p=i(10),l=o(),u=process.env.port||process.env.PORT||3007;l.use(o.json({verify:(t,e,i,o)=>{t.rawBody=i.toString()}})),l.use(o.urlencoded({extended:!0})),l.set("views",a.join(__dirname,"/")),l.use(s("tiny")),l.use("/scripts",o.static(a.join(__dirname,"web/scripts"))),l.use("/assets",o.static(a.join(__dirname,"web/assets"))),l.use(r.MsTeamsApiRouter(p)),l.use(r.MsTeamsPageRouter({root:a.join(__dirname,"web/"),components:p})),l.use("/",o.static(a.join(__dirname,"web/"),{index:"index.html"})),l.set("port",u),n.createServer(l).listen(u,()=>{c("Server running on "+u)})},function(t,e){t.exports=require("express")},function(t,e){t.exports=require("http")},function(t,e){t.exports=require("path")},function(t,e){t.exports=require("morgan")},function(t,e){t.exports=require("dotenv")},function(t,e,i){"use strict";function o(t){for(var i in t)e.hasOwnProperty(i)||(e[i]=t[i])}Object.defineProperty(e,"__esModule",{value:!0}),e.nonce={},o(i(11)),o(i(12))},function(t,e,i){"use strict";var o=this&&this.__decorate||function(t,e,i,o){var n,a=arguments.length,s=a<3?e:null===o?o=Object.getOwnPropertyDescriptor(e,i):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)s=Reflect.decorate(t,e,i,o);else for(var r=t.length-1;r>=0;r--)(n=t[r])&&(s=(a<3?n(s):a>3?n(e,i,s):n(e,i))||s);return a>3&&s&&Object.defineProperty(e,i,s),s};Object.defineProperty(e,"__esModule",{value:!0});const n=i(0);let a=class{};a=o([n.PreventIframe("/acPrototypeTab/index.html")],a),e.AcPrototypeTab=a},function(t,e,i){"use strict";var o=this&&this.__decorate||function(t,e,i,o){var n,a=arguments.length,s=a<3?e:null===o?o=Object.getOwnPropertyDescriptor(e,i):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)s=Reflect.decorate(t,e,i,o);else for(var r=t.length-1;r>=0;r--)(n=t[r])&&(s=(a<3?n(s):a>3?n(e,i,s):n(e,i))||s);return a>3&&s&&Object.defineProperty(e,i,s),s},n=this&&this.__awaiter||function(t,e,i,o){return new(i||(i=Promise))((function(n,a){function s(t){try{c(o.next(t))}catch(t){a(t)}}function r(t){try{c(o.throw(t))}catch(t){a(t)}}function c(t){t.done?n(t.value):new i((function(e){e(t.value)})).then(s,r)}c((o=o.apply(t,e||[])).next())}))};Object.defineProperty(e,"__esModule",{value:!0});const a=i(0),s=i(1),r=i(2),c=i(13),p=i(14),l=i(15),u=i(16),d=i(18),m=i(20),y=i(21),h=i(23),g=i(25),f=i(27),v=i(29);s("msteams");let b=class{constructor(t){this.activityProc=new m.TeamsActivityProcessor,this.credentials=new p.MicrosoftAppCredentials(process.env.MICROSOFT_APP_ID||"",process.env.MICROSOFT_APP_PASSWORD||""),this.conversationState=t,this.dialogState=t.createProperty("dialogState"),this.dialogs=new r.DialogSet(this.dialogState),this.dialogs.add(new l.default("help")),p.MicrosoftAppCredentials.trustServiceUrl("https://smba-int.cloudapp.net/teams-int-mock/"),this.activityProc.messageActivityHandler={onMessage:t=>n(this,void 0,void 0,(function*(){const e=m.TeamsContext.from(t);switch(t.activity.type){case c.ActivityTypes.Message:const i=e?e.getActivityTextWithoutMentions().toLowerCase():t.activity.text;if(i.startsWith("hello"))return void(yield t.sendActivity("Oh, hello to you as well!"));if(i.startsWith("help")){const e=yield this.dialogs.createContext(t);yield e.beginDialog("help")}else if(i.includes("card"))try{const e=c.CardFactory.adaptiveCard(u.default);yield t.sendActivity({attachments:[e]})}catch(t){console.log(t)}else console.log(i),yield t.sendActivity("I'm terribly sorry, but my master hasn't trained me to do anything yet...");break;case c.ActivityTypes.Invoke:}return this.conversationState.saveChanges(t)}))},this.activityProc.conversationUpdateActivityHandler={onConversationUpdateActivity:t=>n(this,void 0,void 0,(function*(){if(t.activity.membersAdded&&0!==t.activity.membersAdded.length)for(const e in t.activity.membersAdded)if(t.activity.membersAdded[e].id===t.activity.recipient.id){const e=c.CardFactory.adaptiveCard(u.default);yield t.sendActivity({attachments:[e]})}}))},this.activityProc.messageReactionActivityHandler={onMessageReaction:t=>n(this,void 0,void 0,(function*(){const e=t.activity.reactionsAdded;e&&e[0]&&(yield t.sendActivity({textFormat:"xml",text:`That was an interesting reaction (<b>${e[0].type}</b>)`}))}))},this.activityProc.invokeActivityHandler={onInvoke:t=>n(this,void 0,void 0,(function*(){const e=t,i=c.CardFactory.adaptiveCard(u.default),o=c.CardFactory.adaptiveCard(y.default),n=c.CardFactory.adaptiveCard(h.default),a=c.CardFactory.adaptiveCard(g.default),s=c.CardFactory.adaptiveCard(d.default),r=c.CardFactory.adaptiveCard(f.default),p=c.CardFactory.adaptiveCard(v.default);let l;a.content.$data={creator:{name:e.activity.name,profileImage:"https://randomuser.me/api/portraits/women/32.jpg"}};const m={tab:{type:"continue",value:{cards:[{card:n.content},{card:a.content},{card:o.content}]}}},b={tab:{type:"continue",value:{cards:[{card:i.content},{card:r.content},{card:s.content}]}}},w={tab:{type:"continue",value:{cards:[{card:p.content},{card:n.content},{card:a.content},{card:o.content}]}}};switch(e.activity.name){case"task/fetch":l={task:{type:"continue",value:{height:"medium",width:"medium",title:"task",card:s}}};break;case"task/submit":l={task:{type:"continue",value:w}};break;case"tab/submit":l=w;break;case"tab/fetch":default:l="workday"===e.activity.value.tabContext.tabEntityId?m:b}return{status:200,body:l}}))}}onTurn(t){return n(this,void 0,void 0,(function*(){yield this.activityProc.processIncomingActivity(t)}))}};b=o([a.BotDeclaration("/api/messages",new c.MemoryStorage,process.env.MICROSOFT_APP_ID,process.env.MICROSOFT_APP_PASSWORD),a.PreventIframe("/acPrototypeBot/acProtoBotTab.html")],b),e.AcPrototypeBot=b},function(t,e){t.exports=require("botbuilder")},function(t,e){t.exports=require("botframework-connector")},function(t,e,i){"use strict";var o=this&&this.__awaiter||function(t,e,i,o){return new(i||(i=Promise))((function(n,a){function s(t){try{c(o.next(t))}catch(t){a(t)}}function r(t){try{c(o.throw(t))}catch(t){a(t)}}function c(t){t.done?n(t.value):new i((function(e){e(t.value)})).then(s,r)}c((o=o.apply(t,e||[])).next())}))};Object.defineProperty(e,"__esModule",{value:!0});const n=i(2);class a extends n.Dialog{constructor(t){super(t)}beginDialog(t,e){return o(this,void 0,void 0,(function*(){return t.context.sendActivity("I'm just a friendly but rather stupid bot, and right now I don't have any valuable help for you!"),yield t.endDialog()}))}}e.default=a},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(17);e.default=o},function(t){t.exports={$schema:"http://adaptivecards.io/schemas/adaptive-card.json",type:"AdaptiveCard",version:"1.0",body:[{type:"TextBlock",spacing:"medium",size:"default",weight:"bolder",text:"More actions ✨",wrap:!0,maxLines:0},{type:"TextBlock",size:"default",isSubtle:!0,text:"Hello, nice to meet you!",wrap:!0,maxLines:0,id:"textToToggle",isVisible:!1},{type:"ColumnSet",isVisible:!1,id:"imagesToToggle",columns:[{type:"Column",items:[{style:"person",type:"Image",url:"https://picsum.photos/300?image=1025",isVisible:!1,id:"imageToToggle",altText:"sample image 1",size:"medium"}]},{type:"Column",items:[{type:"Image",url:"https://picsum.photos/300?image=433",isVisible:!1,id:"imageToToggle2",altText:"sample image 2",size:"medium"}]}]}],actions:[{type:"Action.ToggleVisibility",title:"Get started",targetElements:["textToToggle","imagesToToggle","imageToToggle2"]},{type:"Action.ShowCard",title:"Learn more",card:{type:"AdaptiveCard",body:[{type:"TextBlock",text:"Adaptive Card Tabs Spec"}],actions:[{type:"Action.OpenUrl",title:"Open in browser",url:"https://microsoft.sharepoint.com/:p:/t/ExtensibilityandFundamentals/EWCKez9sB3NEta260kfdMnkBFUu16Z9rnN-gugIqh8D5QQ?e=tBMY9U"}]}},{type:"Action.OpenUrl",title:"acPrototype",url:"https://acprototype.azurewebsites.net"},{type:"Action.Submit",title:"Advanced",data:{msteams:{type:"task/fetch"}}}]}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(19);e.default=o},function(t){t.exports={$schema:"http://adaptivecards.io/schemas/adaptive-card.json",type:"AdaptiveCard",version:"1.0",body:[{type:"Container",items:[{type:"TextBlock",text:"Video Player",weight:"Bolder",size:"Medium"}]},{type:"Container",items:[{type:"TextBlock",text:"Enter the ID of a YouTube video to play in a dialog",wrap:!0},{type:"Input.Text",id:"youTubeVideoId",value:""}]}],actions:[{type:"Action.Submit",title:"Update"}]}},function(t,e){t.exports=require("botbuilder-teams")},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(22);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"Container",bleed:!0,style:"emphasis",items:[{type:"TextBlock",text:"Admin Dashboard",color:"Dark",fontType:"Default",size:"Medium",weight:"Bolder"}],spacing:"None"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Configure the Workday App",wrap:!0},{type:"TextBlock",spacing:"None",text:"Decide what features you want enabled/disabled",isSubtle:!0,wrap:!0,fontType:"Default",weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Configure"}]}]}]},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",wrap:!0,text:"View Diagnostics"},{type:"TextBlock",spacing:"None",text:"Troubleshooting a problem or needing to log out",isSubtle:!0,weight:"Lighter",fontType:"Default",wrap:!0}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"View diagnostics"}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Disconnect from Workday",wrap:!0},{type:"TextBlock",spacing:"None",text:"⚠️ Taking this action will sever all connections between Workday and Microsoft Teams for all of your users.",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Disconnect"}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Innovation Services Agreement",wrap:!0},{type:"TextBlock",spacing:"None",text:"Status: ENABLED\nCredentials: Logan McNeil (CHRO, HRExe, Mgr 4000, MatMgr, ProMgr,TalMgr,VPRept)",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"150px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"Use my credentials"}]}]}],separator:!0,spacing:"Medium"}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(24);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"TextBlock",size:"Medium",weight:"Bolder",text:"Quick actions ✨"},{type:"ActionSet",actions:[{type:"Action.Submit",title:"Take time off"},{type:"Action.Submit",title:"Look up a coworker",data:{msteams:{type:"tab/submit"}}},{type:"Action.Submit",title:"Give feedback",data:{msteams:{type:"tab/submit"}}},{type:"Action.Submit",title:"Ask Workday",data:{msteams:{type:"tab/submit"}}}],separator:!0}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(26);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"Container",bleed:!0,style:"emphasis",items:[{type:"TextBlock",text:"Manager Dashboard | 3 employees",weight:"Bolder",size:"Medium"}]},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/32.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Charlotte Crum",wrap:!0},{type:"TextBlock",spacing:"None",text:"🎉   5 year anniversary this week",isSubtle:!0,wrap:!0,fontType:"Default",weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Small"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/36.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",wrap:!0,text:"Alexa Edwards"},{type:"TextBlock",spacing:"None",text:"📅  Upcoming time off: Mar 20-25, Apr 1",isSubtle:!0,weight:"Lighter",fontType:"Default"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/40.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Yingdan Huang",wrap:!0},{type:"TextBlock",spacing:"None",text:"🎂  Birthday this week",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(28);e.default=o},function(t){t.exports={type:"AdaptiveCard",body:[{type:"Container",bleed:!0,style:"emphasis",items:[{type:"TextBlock",text:"Interview Candidates | 5 candidates",weight:"Bolder",size:"Medium"}]},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/men/21.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Peter Smith",wrap:!0},{type:"TextBlock",spacing:"None",text:"Software Engineer 2 | Bend, Oregon",isSubtle:!0,wrap:!0,fontType:"Default",weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Small"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/31.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",wrap:!0,text:"Marie Beaudoin"},{type:"TextBlock",spacing:"None",text:"Senior Product Manager | Boulder, Colorado",isSubtle:!0,weight:"Lighter",fontType:"Default"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/women/40.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Susan Shammas",wrap:!0},{type:"TextBlock",spacing:"None",text:"Design Manager | Vancouver, British Columbia",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/men/40.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"Aaron Buxton",wrap:!0},{type:"TextBlock",spacing:"None",text:"Senior Software Engineer | Bend, Oregon",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"},{type:"ColumnSet",columns:[{type:"Column",items:[{type:"Image",style:"Person",url:"https://randomuser.me/api/portraits/men/3.jpg",size:"Small"}],width:"auto"},{type:"Column",items:[{type:"TextBlock",weight:"Bolder",text:"John Barry",wrap:!0},{type:"TextBlock",spacing:"None",text:"Design Manager | Seattle, Washington",isSubtle:!0,wrap:!0,weight:"Lighter"}],width:"stretch"},{type:"Column",width:"50px",items:[{type:"ActionSet",actions:[{type:"Action.Submit",title:"..."}]}]}],separator:!0,spacing:"Medium"}],$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2"}},function(t,e,i){"use strict";Object.defineProperty(e,"__esModule",{value:!0});const o=i(30);e.default=o},function(t){t.exports={type:"AdaptiveCard",$schema:"http://adaptivecards.io/schemas/adaptive-card.json",version:"1.2",body:[{type:"Container",items:[{type:"RichTextBlock",inlines:[{type:"TextRun",text:"✅ Success!"}]}],style:"good"}]}}]);
//# sourceMappingURL=server.js.map