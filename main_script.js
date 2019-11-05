// $("#setup").click(() => tryCatch(setup));
// $("#count").click(() => tryCatch(run));
// $("#spell").click(() => tryCatch(replace_wrong_word));
// $("#refresh").click(() => tryCatch(refresh));
//
// let cur_paragraph = 0;
// let cur_word = 0;
// let word_context;
// let current_word;
// let paragraphs;
// let correct_word;
// let temp = {};
// const delimiters = ' ,.:;()!@#$%^&*{}-_+=|?/"' + "'";
// const remove_strategy = [];
// delimiters.split("").forEach((item) => {
//     remove_strategy.push(item);
// });
// const replace_wrong_word = async () => {
//     correct_word = $("#select")
//         .find(":selected")
//         .text();
//     console.log("You Selected: " + correct_word);
//     run();
// };
//
// const getContext = async (term, cur_index) => {
//     const left = 2;
//     const right = 2;
//     const context_words = [];
//
//     const start = Math.max(0, cur_index - 2);
//     for (let i = start; i < cur_index; i++) {
//         context_words.push(term.items[i].text);
//     }
//     for (let i = cur_index + 1; i < cur_index + 3 && i < term.items.length; i++) {
//         context_words.push(term.items[i].text);
//     }
//     return context_words;
// };
//
// const checker = async () => {
//     console.log("Length: " + paragraphs.items.length);
//     for (; cur_paragraph < paragraphs.items.length; cur_paragraph++) {
//         console.log("Cur Para: " + cur_paragraph);
//         console.log("Cur Word:" + cur_word);
//         let paragraph = paragraphs.items[cur_paragraph];
//         await console.log(paragraph);
//         let term;
//         try {
//             term = await paragraph.split(remove_strategy, true, true);
//             await term.load("text,font,style");
//             await word_context.sync();
//             console.log("Done loading");
//         } catch (e) {
//             console.log("Before Word Traverse, found error");
//             console.log(e);
//             return;
//         }
//         console.log("here");
//
//         for (; cur_word < term.items.length; cur_word++) {
//             if (correct_word) {
//                 console.log(term.items.length);
//                 term.items[cur_word].insertText(correct_word, "Replace");
//                 Object.keys(temp.font).forEach((key) => {
//                     term.items[cur_word].font[key] = temp.font[key];
//                 });
//
//                 await word_context.sync();
//                 console.log("inside " + temp.font.color);
//                 correct_word = "";
//
//                 console.log(term.items[cur_word].font.color);
//                 continue;
//             }
//             current_word = term.items[cur_word];
//             const context_words = await getContext(term, cur_word);
//             console.log(current_word.text + "-> " + context_words);
//             if (cur_word % 2 == 0) {
//                 $("#select")
//                     .find("option")
//                     .remove();
//                 const suggestions = await get_suggestion(current_word.text, context_words);
//                 temp.font = JSON.stringify(term.items[cur_word].font);
//                 temp.font = JSON.parse(temp.font);
//                 // term.items[cur_word].select(Word.SelectionMode.select);
//                 term.items[cur_word].font.color = "red";
//                 term.items[cur_word].font.underline = Word.UnderlineType.waveHeavy;
//                 await word_context.sync();
//                 suggestions.forEach((suggestion) => {
//                     $("#select").append("<option value=" + suggestion + ">" + suggestion + "</option>");
//                 });
//                 return;
//             }
//         }
//         cur_word = 0;
//     }
// };
//
// const checker2 = async () => {
//     for (; cur_paragraph < paragraphs.items.length; cur_paragraph++) {
//         console.log("Cur Para: " + cur_paragraph);
//         console.log("Cur Word:" + cur_word);
//         let paragraph = paragraphs.items[cur_paragraph];
//         await console.log(paragraph);
//         let term;
//         try {
//             term = await paragraph.split(remove_strategy, true, true);
//             await term.load();
//             await word_context.sync();
//             console.log("Done loading");
//         } catch (e) {
//             console.log("Before Word Traverse, found error");
//             console.log(e);
//             return;
//         }
//         console.log("here");
//
//         for (; cur_word < term.items.length; cur_word++) {
//             if (correct_word) {
//                 console.log(term.items.length);
//                 const cc = term.items[cur_word].contentControls;
//                 await cc.load();
//                 await word_context.sync();
//                 console.log("CC " + JSON.stringify(cc));
//                 term.items[cur_word].insertText(correct_word, "Replace");
//                 // Object.keys(temp.font).forEach((key) => {
//                 //   term.items[cur_word].font[key] = temp.font[key];
//                 // });
//
//                 await word_context.sync();
//                 // console.log("inside " + temp.font.color);
//                 correct_word = false;
//
//                 // console.log(term.items[cur_word].font.color);
//                 continue;
//             }
//             current_word = term.items[cur_word];
//             const context_words = await getContext(term, cur_word);
//             console.log(current_word.text + "-> " + context_words);
//             if (cur_word % 2 == 0) {
//                 $("#select")
//                     .find("option")
//                     .remove();
//                 const suggestions = await get_suggestion(current_word.text, context_words);
//                 let contentC = term.items[cur_word].properties;
//                 contentC.load();
//                 await word_context.sync();
//                 console.log("Custom Properties " + JSON.stringify(contentC, 0, 2));
//                 // term.items[cur_word].select(Word.SelectionMode.select);
//                 term.items[cur_word].font.color = "red";
//                 term.items[cur_word].font.underline = Word.UnderlineType.waveHeavy;
//                 await word_context.sync();
//                 suggestions.forEach((suggestion) => {
//                     $("#select").append("<option value=" + suggestion + ">" + suggestion + "</option>");
//                 });
//                 return;
//             }
//         }
//         cur_word = 0;
//     }
// };
//
// const get_suggestion = async (word, contexts) => {
//     return [word + "1", word + "2", word + "3"];
// };
//
// const refresh = async () => {
//     cur_paragraph = 0;
//     cur_word = 0;
//     correct_word = false;
//     current_word = "";
// };
//
// async function run() {
//     await Word.run(async (context) => {
//         word_context = context;
//         paragraphs = word_context.document.body.paragraphs;
//         paragraphs.load();
//         await word_context.sync();
//         await checker();
//         await word_context.sync();
//     });
// }
//
// async function continue_run() {
//     await checker();
//     await word_context.sync();
// }
//
// async function setup() {
//     await Word.run(async (context) => {
//         context.document.body.clear();
//         context.document.body.insertParagraph(
//             "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
//             "Start"
//         );
//         context.document.body.insertParagraph(
//             "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.",
//             "End"
//         );
//     });
// }
//
// /** Default helper for invoking an action and handling errors. */
// async function tryCatch(callback) {
//     try {
//         await callback();
//     } catch (error) {
//         // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
//         console.error(error);
//     }
// }
