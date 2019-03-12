// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const osLocale = require("os-locale");
const {ActivityTypes, CardFactory, ActionTypes} = require("botbuilder");
const {ChoicePrompt, DialogSet, OAuthPrompt, WaterfallDialog} = require("botbuilder-dialogs");
const Graph = require("./graph");
const tools = require('./gsuite');

const Language = require('./Resources/Language.json');

// Name of the dialog state used in the constructor.
const DIALOG_STATE_PROPERTY = "dialogState";

// Names of the prompts the bot uses.
const OAUTH_PROMPT = "oAuth_prompt";
const CONFIRM_PROMPT = "confirm_prompt";

// Name of the WaterfallDialog the bot uses.
const AUTH_DIALOG = "auth_dialog";

// The connection name here must match the one from
// your Bot Channels Registration on the settings blade in Azure.
const CONNECTION_NAME = "DENSTU_ISOBAR_AUTH";

let detected_lang;
let lang;
// (async () => {
//   detected_lang  = await osLocale();
//   if(detected_lang.toLowerCase().includes('jp') || detected_lang.toLowerCase().includes('japan'))
//     lang = 'jp';
//   else
//     lang = 'en';
//   console.log("++++ ", detected_lang)
// })();

let userProfile;
lang = 'en';  // JUST FOR TEST

// Create the settings for the OAuthPrompt.
const OAUTH_SETTINGS = {
  connectionName: CONNECTION_NAME,
  title: Language[lang].signIn,
  // title: "Sign In",
  text: Language[lang].signInTitle,
  // text: "Please Sign In",
  timeout: 300000 // User has 5 minutes to log in.
};


// Import Microsoft.Recognizers.Text
const Recognizers = require("@microsoft/recognizers-text-suite");
// const DateTimeRecognizers = require('@microsoft/recognizers-text-date-time');

// Adaptive Card Content (IntroCard)
const IntroCard = require("./Resources/IntroCard.json");

// The accessor name for the conversation data and user profile state property accessors.
const CONVERSATION_DATA_PROPERTY = "conversationData";
const USER_PROFILE_PROPERTY = "userProfile";
const question = {
  none: "none",
  name: "name",
  attendee: "attendee",
  duration: "duration",
  title: "title",
  space: "space",
  date: "date",
  room: "room"
};

let meetingTime = {};
let typeSpace = "";
const allEmployees = [];
const employeesList = allEmployees;
let employeesToShow = [];
let availableDateTimeList = [];
let availableRoomList = [];
// let MeetingRooms = [];
let object = {};
let token = '';

class MyBot {
  constructor(conversationState, userState) {
    // Create the state property accessors for the conversation data and user profile.
    this.conversationData = conversationState.createProperty(
      CONVERSATION_DATA_PROPERTY
    );
    this.userProfile = userState.createProperty(USER_PROFILE_PROPERTY);

    // The state management objects for the conversation and user state.
    this.conversationState = conversationState;
    this.userState = userState;

    // Create a new state accessor property.
    // See https://aka.ms/about-bot-state-accessors to learn more about bot state and state accessors.
    this.dialogState = this.conversationState.createProperty(
      DIALOG_STATE_PROPERTY
    );
    this.dialogs = new DialogSet(this.dialogState);

    // Add prompts that will be used by the bot.
    this.dialogs.add(new ChoicePrompt(CONFIRM_PROMPT));
    this.dialogs.add(new OAuthPrompt(OAUTH_PROMPT, OAUTH_SETTINGS));

    // The WaterfallDialog that controls the flow of the conversation.
    this.dialogs.add(
      new WaterfallDialog(AUTH_DIALOG, [
        this.oauthPrompt,
        this.loginResults,
        this.displayToken
      ])
    );
  }

  /**
   * Waterfall step that prompts the user to login if they have not already or their token has expired.
   * @param {WaterfallStepContext} step
   */
  async oauthPrompt(step) {
    return await step.prompt(OAUTH_PROMPT);
  }

  /**
   * Waterfall step that informs the user that they are logged in and asks
   * the user if they would like to see their token via a prompt
   * @param {WaterfallStepContext} step
   */
  async loginResults(step) {
    let tokenResponse = step.result;
    if (tokenResponse != null) {
      await step.context.sendActivity(Language[lang].loggedInSuccess);
      token = tokenResponse.token;
      // console.log(token);
      return await step.endDialog();
      // return await step.prompt(CONFIRM_PROMPT, 'Do you want to view your token?', ['yes', 'no']);
    }

    // Something went wrong, inform the user they were not logged in
    await step.context.sendActivity(Language[lang].loggedInFail);
    return await step.endDialog();
  }

  /**
   *
   * Waterfall step that will display the user's token. If the user's token is expired
   * or they are not logged in this will prompt them to log in first.
   * @param {WaterfallStepContext} step
   */
  async displayToken(step) {
    const result = step.result.value;
    if (result === "yes") {
      // Call the prompt again because we need the token. The reasons for this are:
      // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
      // about refreshing it. We can always just call the prompt again to get the token.
      // 2. We never know how long it will take a user to respond. By the time the
      // user responds the token may have expired. The user would then be prompted to login again.
      //
      // There is no reason to store the token locally in the bot because we can always just call
      // the OAuth prompt to get the token or get a new token if needed.
      let prompt = await step.prompt(OAUTH_PROMPT);
      let tokenResponse = prompt.result;
      if (tokenResponse != null) {
        await step.context.sendActivity(`${Language[lang].tokenDisplay} ${tokenResponse.token}`);
        await step.context.sendActivity(Language[lang].HELP_TEXT);
        return await step.endDialog();
      }
    }

    await step.context.sendActivity(Language[lang].HELP_TEXT);
    return await step.endDialog();
  }

  /**
   *
   * @param {TurnContext} on turn context object.
   */
  // The bot's turn handler
  async onTurn(turnContext) {
    // This bot listens for message activities.
    if (turnContext.activity.type === ActivityTypes.Message) {
      // Create a dialog context object.
      const dc = await this.dialogs.createContext(turnContext);
      const text = (turnContext.activity.text).toLowerCase();

      // const userProfile = await this.userProfile.get(turnContext, {});
      const conversationData = await this.conversationData.get(turnContext, {
        lastQuestionAsked: question.none
      });

      // Create an array with the valid options.
      const validCommands = ["logout", "help", "restart", "ログアウト", "ヘルプ", "やり直す"];
      await dc.continueDialog();

      // If the user asks for help, send a message to them informing them of the operations they can perform.
      if (validCommands.includes(text)) {
        if (text === "help" || text === "ヘルプ") {
          await turnContext.sendActivity(Language[lang].HELP_TEXT);
        }
        // Log the user out
        if (text === "logout" || text === "ログアウト") {
          let botAdapter = turnContext.adapter;
          await botAdapter.signOutUser(turnContext, CONNECTION_NAME);
          await turnContext.sendActivity(Language[lang].loggedOutSuccess);
          conversationData.lastQuestionAsked = question.none;
          token = "";
          await turnContext.sendActivity(Language[lang].HELP_TEXT);
        }
        // Reset all the input stored
        if (text === "restart" || text === "やり直す") {
          conversationData.lastQuestionAsked = question.none;
          await turnContext.sendActivity(Language[lang].RESTART_TEXT);
        }
      } else {
        if (!turnContext.responded && !token) {
          await dc.beginDialog(AUTH_DIALOG);
        } else {
          userProfile = await this.userProfile.get(turnContext, {});
          const conversationData = await this.conversationData.get(
            turnContext,
            { lastQuestionAsked: question.none }
          );

          await MyBot.fillOutMeetingInfo(
            conversationData,
            userProfile,
            turnContext
          );

          // Update state and save changes.
          await this.userProfile.set(turnContext, userProfile);
          await this.userState.saveChanges(turnContext);

          await this.conversationData.set(turnContext, conversationData);
          await this.conversationState.saveChanges(turnContext);
        }
      }
    } else if (turnContext.activity.type === ActivityTypes.ConversationUpdate) {
        let botAdapter = turnContext.adapter;
        await botAdapter.signOutUser(turnContext, CONNECTION_NAME);
        // await turnContext.sendActivity(Language[lang].loggedOutSuccess);
        // conversationData.lastQuestionAsked = question.none;
        token = "";
        // await turnContext.sendActivity(Language[lang].HELP_TEXT);
      await this.sendWelcomeMessage(turnContext);
    } else if (
      turnContext.activity.type === ActivityTypes.Invoke ||
      turnContext.activity.type === ActivityTypes.Event
    ) {
      // This handles the MS Teams Invoke Activity sent when magic code is not used.
      // See: https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/auth-oauth-card#getting-started-with-oauthcard-in-teams
      // The Teams manifest schema is found here: https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema
      // It also handles the Event Activity sent from the emulator when the magic code is not used.
      // See: https://blog.botframework.com/2018/08/28/testing-authentication-to-your-bot-using-the-bot-framework-emulator/
      const dc = await this.dialogs.createContext(turnContext);
      await dc.continueDialog();
      if (!turnContext.responded) {
        await dc.beginDialog(AUTH_DIALOG);
      }
    } else {
      await turnContext.sendActivity(
        `[${turnContext.activity.type} event detected.]`
      );
    }
    // Update the conversation state before ending the turn.
   await this.conversationState.saveChanges(turnContext);
  }

  // Manages conversation flow for filling out the user's profile
  static async fillOutMeetingInfo(conversationData, userProfile, turnContext) {
    // Check I'll start arranging meeting with, translation has to be dynamic
    const input = turnContext.activity.text;
    let result;
    let splitResults;

    if (input === "External" || input === "外部" || input === "Internal" || input === "内部") {
      conversationData.lastQuestionAsked = question.title;
    }
    switch (conversationData.lastQuestionAsked) {
      // If we're just starting off, we haven't asked the user for any information yet.
      // Ask the user for their intend and update the conversation flag.
      case question.none:
        // build buttons to display.
        const buttons = [
          {
            type: ActionTypes.ImBack,
            // title: msg[0][lang],
            title: Language[lang].startToBookMeetingRoom,
            // value: "Start to book meeting room"
            value: Language[lang].startToBookMeetingRoom
          },
          { type: ActionTypes.ImBack,
            // title: msg[5][lang],
            title: Language[lang].logOut,
            // value: "logout"
            value: Language[lang].logOut
          }
        ];

        // construct hero card.
        const card = CardFactory.heroCard(
            Language[lang].intendCardTitle,
            // msg[1][lang],
            undefined, buttons, {});
        // add card to Activity.
        const reply = { type: ActivityTypes.Message };
        reply.attachments = [card];
        // Send hero card to the user.
        await turnContext.sendActivity(reply);
        conversationData.lastQuestionAsked = question.name;
        break;
      // If we last asked for their intend, record their response, confirm that we got it.
      // Ask the user for the attendees and update the conversation flag.
      case question.name:
        let userList = await Graph.user_list(token);
        userList.forEach(function(obj) {
          allEmployees.push(obj.displayName.trim().toLowerCase());
        });
        result = this.validateInitialInput(input);
        employeesToShow = this.getRandom(employeesList, 3);
        // employeesToShow.push("other");
        employeesToShow.push(Language[lang].other);
        if (result.success) {
          await turnContext.sendActivity("OK. " + Language[lang].startToBookMeetingRoom);
          await this.sendSuggestedAttendees(turnContext);
          userProfile.name = [];
          userProfile.email = [];
          conversationData.lastQuestionAsked = question.attendee;
          break;
        } else {
          // If we couldn't interpret their input, ask them for it again.
          // Don't update the conversation flag, so that we repeat this step.
          // await turnContext.sendActivity(result.message || msg[2][lang]);
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand);
          break;
        }
      // If we last asked for attendees, record their response, confirm that we got it.
      // Ask the user for duration of the meeting and update the conversation flag.
      case question.attendee:
        if (input.trim().toLowerCase() !== "no, there isn't" && input !== "いません。") {
          splitResults = input.split(",");
        } else {
          if (lang === "jp") {
            await turnContext.sendActivity(`${userProfile.name} ${Language[lang].arrangeAttendee}`);
          } else {
            await turnContext.sendActivity(`${Language.en.arrangeAttendee} ${userProfile.name}.`);
          }
          // await turnContext.sendActivity(Language[lang].arrangeAttendee);
          await turnContext.sendActivity(Language[lang].durationInput);
          conversationData.lastQuestionAsked = question.duration;
          splitResults = "";
          break;
        }
        if (input.trim().toLowerCase() === "other" || input === "その他") {
          await turnContext.sendActivity(`${Language[lang].whoAreAttendees}
                            ${Language[lang].ableToFindWithNameEmail}`);
          break;
        }

        if (splitResults.length >= 2) {
          for (let i = 0; i < splitResults.length; i++) {
            let splitResult = splitResults[i].trim();
            result = await Graph.searchUser(token, splitResult);
            if (result.status) {
              userProfile.name.push(result.data.displayName);
              userProfile.email.push(result.data.mail);
            } else {
              console.log(splitResult + " declined");
            }
          }
          if (lang === "jp") {
            await turnContext.sendActivity(`${userProfile.name} ${Language.en.arrangeAttendee}`);
          } else {
            await turnContext.sendActivity(`${Language.en.arrangeAttendee} ${userProfile.name}.`);
          }
          await turnContext.sendActivity(Language[lang].durationInput);
          conversationData.lastQuestionAsked = question.duration;
          break;
        }
        result = await Graph.searchUser(token, input);
        if (result.status) {
          userProfile.name.push(result.data.displayName);
          userProfile.email.push(result.data.mail);
          let index = employeesList.indexOf(
            result.data.displayName.toLowerCase().trim()
          );
          if (index > -1) {
            employeesList.splice(index, 1);
            if (employeesList.length > 3) {
              employeesToShow = this.getRandom(employeesList, 2);
              // employeesToShow.push("Other");
              employeesToShow.push(Language[lang].other);
            }
          }
          await this.sendAnotherSuggestedAttendees(turnContext);
          break;
        } else {
          // If we couldn't interpret their input, ask them for it again.
          // Don't update the conversation flag, so that we repeat this step.
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand + Language[lang].invalidEmployee);
          await this.sendSuggestedAttendees(turnContext);
          break;
        }
      // If we last ask for their duration, record their response, confirm that we got it.
      //  Ask them for their for meeting title and update the conversation flag.
      case question.duration:
        result = this.validateDuration(input);
        if (result.success) {
          userProfile.duration = result.duration;
          if (lang === "jp") {
            // await turnContext.sendActivity(`会議を ${userProfile.duration} 分間手配します。`);
            await turnContext.sendActivity(`${userProfile.duration} ${Language[lang].arrangeDuration}`);
          } else {
            await turnContext.sendActivity(`${Language[lang].arrangeDuration} ${userProfile.duration} mns.`);
          }
          // await turnContext.send(`${Language.en.arrangeDuration} ${userProfile.duration} mns.`);
          await turnContext.sendActivity(Language[lang].meetingNameInput);
          conversationData.lastQuestionAsked = question.space;
          break;
        } else {
          // If we couldn't interpret their input, ask them for it again.
          // Don't update the conversation flag, so that we repeat this step.
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand);
          break;
        }
      // If we last asked for a meeting title, record their response, confirm that we got it,
      // Ask them for their preference date, and update the conversation flag.
      case question.space:
        result = this.validateTitle(input);
        meetingTime = await Graph.findMeetingTimes(token,
          { attendees: userProfile.email, duration: userProfile.duration });

      if (result.success) {
          userProfile.title = result.title;
          if (lang === "jp") {
            await turnContext.sendActivity(
              `${userProfile.title} としてあなたのために会議を手配します。`
            );
          } else {
            await turnContext.sendActivity(
              `${Language.en.arrangeMeetingName} ${userProfile.title}.`
            );
          }
          // build buttons to display.
          const buttons = [
            {
              type: ActionTypes.ImBack,
              // title: "External (" + meetingTime.external.length + ")",
              title: `${Language[lang].external} (${meetingTime.external.count})`,
              // value: "External (" + meetingTime.external.count + ")"
              value: `${Language[lang].external} (${meetingTime.external.count})`
            },
            {
              type: ActionTypes.ImBack,
              // title: "Internal (" + meetingTime.internal.length + ")",
              title: `${Language[lang].internal} (${meetingTime.internal.length})`,
              // value: "Internal (" + meetingTime.internal.length + ")"
              value: `${Language[lang].internal} (${meetingTime.internal.length})`
            }
          ];

          // construct hero card.
        // const card = CardFactory.heroCard(msg[13][lang], undefined, buttons, {});
        const card = CardFactory.heroCard(Language[lang].categorizedRoomTypeSuggestionTitle, undefined, buttons, {});
        // add card to Activity.
        const reply = { type: ActivityTypes.Message };
        reply.attachments = [card];
        // Send hero card to the user.
        await turnContext.sendActivity(reply);
        conversationData.lastQuestionAsked = question.title;
        break;
      } else {
          await turnContext.sendActivity(result.message || "I'm sorry. I didn't understand that.");
          await this.sendSuggestedDates(turnContext);
          break;
      }
      // If we last ask for a date, record their response, confirm that we got it.
      // Ask them for their preference room, and update the conversation flag.
      case question.title:
        // console.log(input);
        availableDateTimeList = [];
        result = this.validateTitle(input);
        result.success = true;
        if (input.includes("External") || input.includes("外部")) {
          typeSpace = "external";
          // console.log("########", typeSpace);
          // availableDateTimeList.push("Internal");
          availableDateTimeList.push(Language[lang].internal);
          meetingTime.external.data.forEach(function(obj) {
            availableDateTimeList.push(
              obj.time.start.slice(0, 10) +
                " " +
                obj.time.start.slice(11, 16) +
                "-" +
                obj.time.end.slice(11, 16)
            );
          });

        } else if (input.includes("Internal") || input.includes("内部")) {
          typeSpace = "internal";
          // availableDateTimeList.push("External");
          availableDateTimeList.push(Language[lang].external);
          meetingTime.internal.forEach(function(obj) {
            availableDateTimeList.push(
              obj.time.start.slice(0, 10) +
                " " +
                obj.time.start.slice(11, 16) +
                "-" +
                obj.time.end.slice(11, 16)
            );
          });
        }

        availableDateTimeList =  this.removeDups(availableDateTimeList);

        // availableDateTimeList.push("Other");
          availableDateTimeList.push(Language[lang].other);

        if (result.success) {
          userProfile.category = result.title;

          // await turnContext.sendActivity('Which do you like?');
          await this.sendSuggestedDates(turnContext);
          //var re = MessageFactory.suggestedActions(['Red', 'Yellow', 'Blue'], 'What is the best color?');
          conversationData.lastQuestionAsked = question.date;

          break;
        } else {
          // await turnContext.sendActivity(result.message || msg[2][lang]);
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand);
          await this.sendSuggestedDates(turnContext);
          break;
        }

      case question.date:
        // console.log("typeSpaceeeeee", typeSpace);
        let index = availableDateTimeList.indexOf(input);
        if (input.trim().toLowerCase() === "other" || input === Language[lang].other) {
          typeSpace='other';
          // console.log("Other Date & Time");
          await turnContext.sendActivity(`Please specify the date & time`);
          break;
        }
        result = this.validateDate(input);
        let findMeetingRooms = {};
        // console.log("$$$$$$ ", availableRoomList)
        if (result.success) {
          if (typeSpace === "external") {
            findMeetingRooms = meetingTime.external.data[index - 1];
            let op = this.findExternalRoomByDate(input, meetingTime.external.data);
            availableRoomList = [];
            availableRoomList = op.rooms;

            userProfile.date = result.date;
            userProfile.start = op.start;
            userProfile.end = op.end;
          } else if(typeSpace === "internal") {
            findMeetingRooms = await Graph.forceMeetingTimes(token, {
              date: result.other[0].text.split(" ")[0],
              time: result.other[0].text.split(" ")[1].split("-")[0],
              duration: userProfile.duration
            });
            // console.log("^^^^ ", findMeetingRooms)
            availableRoomList = [];
            findMeetingRooms.locations.forEach(function(obj) {
              object[obj.displayName.trim().toLowerCase()] =
                obj.locationEmailAddress;
              availableRoomList.push(obj.displayName);
            });
            userProfile.date = result.date;
            userProfile.start = findMeetingRooms.time.start;
            userProfile.end = findMeetingRooms.time.end;
          }

          // console.log("Meeting Room Object: " + MeetingRooms);
          // findMeetingRooms.forEach(function(obj) {
          //     availableRoomList.push(obj.displayName);
          // });
          else{
            console.log("!!!@@@@", result.other[0].text.split(" ")[0], "   ", result.other[0].text.split(" ")[1]);
            findMeetingRooms = await Graph.forceMeetingTimes(token, {
              date: result.other[0].text.split(" ")[0],
              time: result.other[0].text.split(" ")[1],
              duration: userProfile.duration
            });
            console.log("^^^^", findMeetingRooms);
            availableRoomList = [];
            findMeetingRooms.locations.forEach(function (obj) {
              object[obj.displayName.trim().toLowerCase()] = obj.locationEmailAddress;
              availableRoomList.push(obj.displayName);
            });
            let external = await tools.getGsuiteRooms([findMeetingRooms], userProfile.name.length);
            external.data.forEach(function (obj) {
              object[obj.location[0].displayName.trim().toLowerCase()] =
                  obj.location[0].locationEmailAddress;
              availableRoomList.push(obj.location[0].displayName);
            });
            userProfile.date = result.date;
            userProfile.start = findMeetingRooms.time.start;
            userProfile.end = findMeetingRooms.time.end;

            if (lang === "jp") {
              await turnContext.sendActivity(
                  `あなたの会議は ${userProfile.date} のためのスケジュールです。`
              );
            } else {
              await turnContext.sendActivity(
                  `Your meeting is schedule for ${userProfile.date}. ${result.other[0].text.split(" ")[1]}`
              );
            }
            await this.sendSuggestedRooms(turnContext);
            conversationData.lastQuestionAsked = question.room;
          }

          if (typeSpace !== 'other') {
            if (lang === "jp") {
              await turnContext.sendActivity(
                  `あなたの会議は ${userProfile.date} のためのスケジュールです。`
              );
            } else {
              await turnContext.sendActivity(
                  `Your meeting is schedule for ${userProfile.date} ${result.other[0].text.split(" ")[1]}`
              );
            }
            await this.sendSuggestedRooms(turnContext);
            conversationData.lastQuestionAsked = question.room;
          }
          break;
        } else {
          // If we couldn't interpret their input, ask them for it again.
          // Don't update the conversation flag, so that we repeat this step.
          // await turnContext.sendActivity(result.message || msg[2][lang]);
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand);
          break;
        }
      // If we last ask for a room, record their response, confirm that we got it.
      // Confirm that the process is completed., and update the conversation flag.
      case question.room:
        let data;
        result = this.validateRoom(input);

        if (result.success) {
          userProfile.room = result.room;
          userProfile.roomAdd = object[result.room.trim().toLowerCase()];
          data = {
            subject: userProfile.title,
            start: userProfile.start,
            end: userProfile.end,
            category: userProfile.category,
            location: {
              displayName: userProfile.room,
              locationEmailAddress: userProfile.roomAdd
            },
            attendees: []
          };
          for (let i = 0; i < userProfile.name.length; i++) {
            let obj = {
              emailAddress: {
                address: userProfile.email[i],
                name: userProfile.name[i]
              },
              type: "required"
            };
            data["attendees"].push(obj);
          }
          let room = {
            emailAddress: {
              address: userProfile.roomAdd,
              name: userProfile.room
            }
          };
          if(typeSpace === "internal")
            data["attendees"].push(room);
          // console.log(data);

          // console.log('^^^^^^ ' + await Graph.createEvent(token, data, typeSpace));
          // let findMeetingRooms = await Graph.forceMeetingTimes(token, {
          //     date: new Date(result.date).toISOString().slice(0, 10), time: result.time, duration: userProfile.duration });
          // console.log(findMeetingRooms);
          // findMeetingRooms.locations.forEach(function(obj) {
          //     availableRoomList.push(obj.displayName);
          // });
          let v = {
            jp:
              "会議室が予約されました。私はこの会議を出席者全員のカレンダーにスケジュールしました。",
            en:
              "Meeting room has been booked. I’ve scheduled this meeting to all of attendees's calendar."
          };

          await turnContext.sendActivity(v[lang]);
          if (lang === "jp") {
            await turnContext.sendActivity(`会議の詳細: 
              タイトル:          ${userProfile.title}
              開始:          ${userProfile.start}
              終わり:            ${userProfile.end}
              日付:           ${userProfile.date}
              期間:       ${userProfile.duration}
              参加者:      ${userProfile.name}
              電子メール:         ${userProfile.email}
              ルーム:           ${userProfile.room}`);
            await turnContext.sendActivity(
              "ボットを再実行するために何かを入力。"
            );
          } else {
            await turnContext.sendActivity(`
**Meeting Detail** 
* Title:          ${userProfile.title}
* Start:          ${userProfile.start}
* End:            ${userProfile.end}
* Date:           ${userProfile.date}
* Duration:       ${userProfile.duration}
* Attendees:      ${userProfile.name}
* Emails:         ${userProfile.email}
* Room:           ${userProfile.room}`
            );
            await turnContext.sendActivity("Type anything to run the bot again.");
          }
          conversationData.lastQuestionAsked = question.none;
          userProfile = {};
          break;
        } else {
          // await turnContext.sendActivity(result.message || msg[2][lang]);
          await turnContext.sendActivity(result.message || Language[lang].doNotUnderstand);
          break;
        }
    }
  }

  // *****Validation***** //
  static removeDups(names) {
    let unique = {};
    names.forEach(function(i) {
      if(!unique[i]) {
        unique[i] = true;
      }
    });
    return Object.keys(unique);
  }

  static findExternalRoomByDate(date, rooms){
    let outputs = [];
    let start, end;
    rooms.forEach(function(room) {
        let combinedDate = room.time.start.slice(0, 10) +
        " " +
        room.time.start.slice(11, 16) +
        "-" +
        room.time.end.slice(11, 16);

        if(combinedDate === date){
          object[room.location[0].displayName.trim().toLowerCase()] =
          room.location[0].locationEmailAddress;
          outputs.push(room.location[0].displayName);
          start = room.time.start;
          end = room.time.end;
        }
    });
    return {'rooms':outputs, 'start':start, 'end':end};
  }

  static validateInitialInput(input) {
    const initialInput = input && input.trim().toLowerCase();
    return initialInput === "start to book meeting room" || initialInput === "会議室を予約する"
        ? { success: true, initialInput: initialInput } : { success: false, message: Language[lang].selectOneOptionAbove };
      // : { success: false, message: "Please select one of the option above." };
  }


  static validateDuration(input) {
    // Try to recognize the input as a number. This work for response such as "twelve" as well as "12".
    try {
      // Attempt to convert the recognizer result to an integer. This work for "a dozen", "twelve", "12" and so on.
      // The recognizer returns a list of potential recognition results, if any.
      const results = Recognizers.recognizeNumber(
        input,
        Recognizers.Culture.English
      );
      let output;
      let rr = {
        'jp':'数字だけが受け入れられます。所要時間は5分から180分です。',
        'en': 'Only number is accepted. Duration can be between 5 minutes to 180 minutes.'
      };
      results.forEach(function(result) {
        // result.resolution is a dictionary, where the "value" entry contains the processed string.
        const value = result.resolution["value"];
        if (value) {
          const duration = parseInt(value);
          if (!isNaN(duration) && duration >= 5 && duration <= 180) {
            output = { success: true, duration: duration };
            // return;
          }
        }
      });
      return (
        output || {
          success: false,
          message: rr[lang]
        }
      );
    } catch (e) {
      return {
        success: false,
        message:
          "I'm sorry, I could not interpret that as an age. Please enter an age between 18 and 120."
      };
    }
  }

  // *****Meeting Name Validation***** //
  static validateTitle(input) {
    const title = input && input.trim();
    return title !== undefined
      ? { success: true, title: title }
      : {
          success: false,
          message: "Please enter a name that contain at least one character."
        };
  }


  // Validates date input. Returns whether validation succeeded and either the parsed and normalized
  // value or a message the bot can use to ask the user again.
  // *****DateTime Validation***** //
  static validateDate(input) {
    try {
      const results = Recognizers.recognizeDateTime(input, Recognizers.Culture.English);
      const now = new Date();
      const earliest = now.getTime() + (60 * 60 * 1000);
      let output;
      results.forEach(function (result) {
        // result.resolution is a dictionary, where the "values" entry contains the processed input.
        result.resolution['values'].forEach(function (resolution) {
          // The processed input contains a "value" entry if it is a date-time value, or "start" and
          // "end" entries if it is a date-time range.
          const datevalue = resolution['value'] || resolution['start'];
          // If only time is given, assume it's for today.
          const datetime = resolution['type'] === 'time'
              ? new Date(`${now.toLocaleDateString()} ${datevalue}`)
              : new Date(datevalue);
          if (datetime && earliest < datetime.getTime()) {
            output = { success: true, date: datetime.toLocaleDateString(), time: datevalue.slice(11,21).toLocaleString(), other:results };
            return;
          }
        });
      });
      return output || { success: false, message: "I'm sorry, please enter a date at least an hour out." };
    } catch (error) {
      return {
        success: false,
        message:
            "I'm sorry, I could not interpret that as an appropriate date. Please enter a date at least an hour out."
      };
    }
  }



  // *****Room Validation***** //
  static validateRoom(input) {
    const room = input && input.trim().toLowerCase();
    try {
      let output;
      availableRoomList.forEach(function(result) {
        const availableRoom = result.trim().toLowerCase();
        if (availableRoom === room) {
          output = { success: true, room: room };
          // return;
        }
      });
      return (
        output || {
          success: false,
          message: "Please enter a name of the room that available here."
        }
      );
    } catch (e) {
      return {
        success: false,
        message:
          "I'm sorry, I could not interpret that as a name. Please enter a valid name."
      };
    }
  }

  // *****Send Suggested List***** //
  static async sendSuggestedAttendees(turnContext) {
      // build buttons to display.
      let employeeButtons = [];
      employeesToShow.forEach(function (employee) {
          employeeButtons.push({
              type: ActionTypes.ImBack,
              title: employee,
              value: employee
          });
      });
    // construct hero card.
    const card = CardFactory.heroCard(Language[lang].attendeeSuggestionTitle, undefined, employeeButtons, {});
    // add card to Activity.
    const reply = { type: ActivityTypes.Message };
    reply.attachments = [card];
    // Send hero card to the user.
    await turnContext.sendActivity(reply);
  }

  // *****Send Another Attendees Suggested List***** //
  static async sendAnotherSuggestedAttendees(turnContext) {
    // employeesToShow.unshift("no, there isn't");
      employeesToShow.unshift(Language[lang].noThereIsNot);
    // build buttons to display.
      let employeeButtons = [];
      employeesToShow.forEach(function (employee) {
          employeeButtons.push({
              type: ActionTypes.ImBack,
              title: employee,
              value: employee
          });
      });
    // construct hero card.
    const card = CardFactory.heroCard(Language[lang].attendeeAnotherSuggestionTitle, undefined, employeeButtons, {});
    // add card to Activity.
    const reply = { type: ActivityTypes.Message };
    reply.attachments = [card];
    // Send hero card to the user.
    await turnContext.sendActivity(reply);
  }

  // *****Send DateTime Suggested List***** //
  static async sendSuggestedDates(turnContext) {
      // build buttons to display.
      let dateButtons = [];
      availableDateTimeList.forEach(function (date) {
          dateButtons.push({
              type: ActionTypes.ImBack,
              title: date,
              value: date
          });
      });
    // construct hero card.
    const card = CardFactory.heroCard(Language[lang].dateTimeSuggestionTitle, undefined, dateButtons, {});
    // add card to Activity.
    const reply = { type: ActivityTypes.Message };
    reply.attachments = [card];
    // Send hero card to the user.
    await turnContext.sendActivity(reply);
  }

  // *****Send Rooms Suggested List***** //
  static async sendSuggestedRooms(turnContext) {
      // build buttons to display.
      let roomButtons = [];
      availableRoomList.forEach(function (room) {
          roomButtons.push({
              type: ActionTypes.ImBack,
              title: room,
              value: room
          });
      });
    // construct hero card.
    const card = CardFactory.heroCard(Language[lang].roomSuggestionTitle, undefined, roomButtons, {});
    // add card to Activity.
    const reply = { type: ActivityTypes.Message };
    reply.attachments = [card];
    // Send hero card to the user.
    await turnContext.sendActivity(reply);
  }

  // Sends welcome messages to conversation members when they join the conversation.
  // Messages are only sent to conversation members who aren't the bot.
  async sendWelcomeMessage(turnContext) {
    const activity = turnContext.activity;
    if (activity.membersAdded) {
      // Iterate over all new members added to the conversation.
      for (const idx in activity.membersAdded) {
        if (activity.membersAdded[idx].id !== activity.recipient.id) {
          await turnContext.sendActivity({
            text: "© 2019 Dentsu Isobar, Inc.",
            attachments: [CardFactory.adaptiveCard(IntroCard[lang])]
          });
        }
      }
    }
  }

  // *****Random Function***** //
  static getRandom(arr, n) {
    let result = new Array(n);
    let len = arr.length;
    let taken = new Array(len);
    if (n > len) {
      throw new RangeError("getRandom: more elements taken than available");
    }
    while (n--) {
      let x = Math.floor(Math.random() * len);
      result[n] = arr[x in taken ? taken[x] : x];
      taken[x] = --len in taken ? taken[len] : len;
    }
    return result;
  }
}
module.exports.MyBot = MyBot;


