const {
    ComponentDialog,
    DialogTurnStatus,
    WaterfallDialog
} = require('botbuilder-dialogs');

const { QnACardBuilder } = require('../utils/qnaCardBuilder');

// Default parameters
const DefaultThreshold = 0.3;
const DefaultTopN = 3;
const DefaultNoAnswer = 'Ups, diese Frage kann ich leider nicht beantworten 😓 Kannst du deine Frage umformulieren? Vielleicht hilft mir das ...';

// Card parameters
const DefaultCardTitle = 'Hast du folgendes gemeint:';
const DefaultCardNoMatchText = 'Nichts davon.';
const DefaultCardNoMatchResponse = 'Danke für dein Feedback.';

// Define value names for values tracked inside the dialogs.
const QnAOptions = 'qnaOptions';
const QnADialogResponseOptions = 'qnaDialogResponseOptions';
const CurrentQuery = 'currentQuery';
const QnAData = 'qnaData';
const QnAContextData = 'qnaContextData';
const PreviousQnAId = 'prevQnAId';

/// QnA Maker dialog.
const QNAMAKER_DIALOG = 'qnamaker-dialog';
const QNAMAKER_BASE_DIALOG = 'qnamaker-base-dailog';

class QnAMakerBaseDialog extends ComponentDialog {
    /**
     * Core logic of QnA Maker dialog.
     * @param {QnAMaker} qnaService A QnAMaker service object.
     */
    constructor(qnaService) {
        super(QNAMAKER_BASE_DIALOG);

        this._qnaMakerService = qnaService;

        this.addDialog(new WaterfallDialog(QNAMAKER_DIALOG, [
            this.callGenerateAnswerAsync.bind(this),
            this.callTrain.bind(this),
            this.checkForMultiTurnPrompt.bind(this),
            this.displayQnAResult.bind(this)
        ]));

        this.initialDialogId = QNAMAKER_DIALOG;
    }

    /**
    * @param {WaterfallStepContext} stepContext contextual information for the current step being executed.
    */
    async callGenerateAnswerAsync(stepContext) {
        // Default QnAMakerOptions
        let qnaMakerOptions = {
            scoreThreshold: DefaultThreshold,
            top: DefaultTopN,
            context: {},
            qnaId: -1
        };

        const dialogOptions = getDialogOptionsValue(stepContext);

        if (dialogOptions[QnAOptions] != null) {
            qnaMakerOptions = dialogOptions[QnAOptions];
            qnaMakerOptions.scoreThreshold = qnaMakerOptions.scoreThreshold ? qnaMakerOptions.scoreThreshold : DefaultThreshold;
            qnaMakerOptions.top = qnaMakerOptions.top ? qnaMakerOptions.top : DefaultThreshold;
        }

        // Storing the context info
        stepContext.values[CurrentQuery] = stepContext.context.activity.text;

        const previousContextData = dialogOptions[QnAContextData];
        const prevQnAId = dialogOptions[PreviousQnAId];

        if (previousContextData != null && prevQnAId != null) {
            if (prevQnAId > 0) {
                qnaMakerOptions.context = {
                    previousQnAId: prevQnAId
                };

                qnaMakerOptions.qnaId = 0;
                if (previousContextData[stepContext.context.activity.text.toLowerCase()] !== null) {
                    qnaMakerOptions.qnaId = previousContextData[stepContext.context.activity.text.toLowerCase()];
                }
            }
        }

        // Calling QnAMaker to get response.
        const response = await this._qnaMakerService.getAnswersRaw(stepContext.context, qnaMakerOptions);

        // Resetting previous query.
        dialogOptions[PreviousQnAId] = -1;
        stepContext.activeDialog.state.options = dialogOptions;

        // Take this value from GetAnswerResponse.
        const isActiveLearningEnabled = response.activeLearningEnabled;

        stepContext.values[QnAData] = response.answers;

        // Check if active learning is enabled.
        if (isActiveLearningEnabled && response.answers.length > 0 && response.answers[0].score <= 0.95) {
            response.answers = this._qnaMakerService.getLowScoreVariation(response.answers);

            const suggestedQuestions = [];
            if (response.answers.length > 1) {
                // Display suggestions card.
                response.answers.forEach(element => {
                    suggestedQuestions.push(element.questions[0]);
                });
                const qnaDialogResponseOptions = dialogOptions[QnADialogResponseOptions];
                const message = QnACardBuilder.GetSuggestionCard(suggestedQuestions, qnaDialogResponseOptions.activeLearningCardTitle, qnaDialogResponseOptions.cardNoMatchText);
                await stepContext.context.sendActivity(message);

                return { status: DialogTurnStatus.waiting };
            }
        }

        const result = [];
        if (response.answers.length > 0) {
            result.push(response.answers[0]);
        }

        stepContext.values[QnAData] = result;

        return await stepContext.next(result);
    }

    /**
    * @param {WaterfallStepContext} stepContext contextual information for the current step being executed.
    */
    async callTrain(stepContext) {
        const trainResponses = stepContext.values[QnAData];
        const currentQuery = stepContext.values[CurrentQuery];

        const reply = stepContext.context.activity.text;

        const dialogOptions = getDialogOptionsValue(stepContext);
        const qnaDialogResponseOptions = dialogOptions[QnADialogResponseOptions];

        if (trainResponses.length > 1) {
            const qnaResults = trainResponses.filter(r => r.questions[0] === reply);

            if (qnaResults.length > 0) {
                stepContext.values[QnAData] = qnaResults;

                const feedbackRecords = {
                    FeedbackRecords: [
                        {
                            UserId: stepContext.context.activity.id,
                            UserQuestion: currentQuery,
                            QnaId: qnaResults[0].id
                        }
                    ]
                };

                // Call Active Learning Train API
                this._qnaMakerService.callTrainAsync(feedbackRecords);

                return await stepContext.next(qnaResults);
            } else if (reply === qnaDialogResponseOptions.cardNoMatchText) {
                await stepContext.context.sendActivity(qnaDialogResponseOptions.cardNoMatchResponse);
                return await stepContext.endDialog();
            } else {
                return await stepContext.replaceDialog(QNAMAKER_DIALOG, stepContext.activeDialog.state.options);
            }
        }

        return await stepContext.next(stepContext.result);
    }

    /**
    * @param {WaterfallStepContext} stepContext contextual information for the current step being executed.
    */
    async checkForMultiTurnPrompt(stepContext) {
        if (stepContext.result != null && stepContext.result.length > 0) {
            // -Check if context is present and prompt exists.
            // -If yes: Add reverse index of prompt display name and its corresponding qna id.
            // -Set PreviousQnAId as answer.Id.
            // -Display card for the prompt.
            // -Wait for the reply.
            // -If no: Skip to next step.

            const answer = stepContext.result[0];

            if (answer.context != null && answer.context.prompts != null && answer.context.prompts.length > 0) {
                const dialogOptions = getDialogOptionsValue(stepContext);

                const previousContextData = {};

                // eslint-disable-next-line no-extra-boolean-cast
                if (!!dialogOptions[QnAContextData]) {
                    previousContextData = dialogOptions[QnAContextData];
                }

                answer.context.prompts.forEach(prompt => {
                    previousContextData[prompt.displayText.toLowerCase()] = prompt.qnaId;
                });

                dialogOptions[QnAContextData] = previousContextData;
                dialogOptions[PreviousQnAId] = answer.id;
                stepContext.activeDialog.state.options = dialogOptions;

                // Get multi-turn prompts card activity.
                const message = QnACardBuilder.GetQnAPromptsCard(answer);
                await stepContext.context.sendActivity(message);

                return { status: DialogTurnStatus.waiting };
            }
        }

        return await stepContext.next(stepContext.result);
    }

    /**
    * @param {WaterfallStepContext} stepContext contextual information for the current step being executed.
    */
    async displayQnAResult(stepContext) {
        const dialogOptions = getDialogOptionsValue(stepContext);
        const qnaDialogResponseOptions = dialogOptions[QnADialogResponseOptions];

        const reply = stepContext.context.activity.text;

        if (reply === qnaDialogResponseOptions.cardNoMatchText) {
            await stepContext.context.sendActivity(qnaDialogResponseOptions.cardNoMatchResponse);
            return await stepContext.endDialog();
        }

        const previousQnAId = dialogOptions[PreviousQnAId];
        if (previousQnAId > 0) {
            return await stepContext.replaceDialog(QNAMAKER_DIALOG, dialogOptions);
        }

        const responses = stepContext.result;
        if (responses != null) {
            if (responses.length > 0) {
                await stepContext.context.sendActivity(responses[0].answer);
            } else {
                await stepContext.context.sendActivity(qnaDialogResponseOptions.noAnswer);
            }
        }

        return await stepContext.endDialog();
    }
}

function getDialogOptionsValue(dialogContext) {
    let dialogOptions = {};

    if (dialogContext.activeDialog.state.options !== null) {
        dialogOptions = dialogContext.activeDialog.state.options;
    }

    return dialogOptions;
}

module.exports.QnAMakerBaseDialog = QnAMakerBaseDialog;
module.exports.QNAMAKER_BASE_DIALOG = QNAMAKER_BASE_DIALOG;
module.exports.DefaultThreshold = DefaultThreshold;
module.exports.DefaultTopN = DefaultTopN;
module.exports.DefaultNoAnswer = DefaultNoAnswer;
module.exports.DefaultCardTitle = DefaultCardTitle;
module.exports.DefaultCardNoMatchText = DefaultCardNoMatchText;
module.exports.DefaultCardNoMatchResponse = DefaultCardNoMatchResponse;
module.exports.QnAOptions = QnAOptions;
module.exports.QnADialogResponseOptions = QnADialogResponseOptions;
