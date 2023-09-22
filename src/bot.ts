import {
  ConversationState,
  UserState,
  TeamsActivityHandler,
  TurnContext,
} from "botbuilder";

export class DialogBot extends TeamsActivityHandler {
  constructor(
    public conversationState: ConversationState,
    public userState: UserState
  ) {
    super();
    this.conversationState = conversationState;
    this.userState = userState;

    this.onMessage(async (context, next) => {
      await context.sendActivity("caught user input");
      await next();
    });
  }

  public async run(context: TurnContext) {
    await super.run(context);
    // Save any state changes. The load happened during the execution of the Dialog.
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }
}
