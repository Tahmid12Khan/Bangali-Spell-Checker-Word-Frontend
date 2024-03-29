var RedactAddin = {};

let cur_paragraph = 0;
let cur_word = 0;
let word_context;
let current_word;
let paragraphs;
let correct_word;
let temp_correct_word;
let error = false;
let temp = {};
let term;
let isRefreshClicked = false;
const delimiters = ' ,.:;()!@#$%^&*{}-_+=|?/"।' + "'";
const remove_strategy = [];
delimiters.split("").forEach((item) => {
    remove_strategy.push(item);
});


async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}


const getContext = async (term, cur_index) => {
    const left = 2;
    const right = 2;
    const context_words = [];

    const start = Math.max(0, cur_index - 2);
    for (let i = start; i < cur_index; i++) {
        context_words.push(term.items[i].text);
    }
    for (let i = cur_index + 1; i < cur_index + 3 && i < term.items.length; i++) {
        context_words.push(term.items[i].text);
    }
    return context_words;
};

const result_for_cur_word = async () => {
    console.log(term.items.length);

    Object.keys(temp.font).forEach((key) => {
        term.items[cur_word].font[key] = temp.font[key];
    });
    if (correct_word) {
        term.items[cur_word].insertText(correct_word, "Replace");
    }
    await word_context.sync();
    console.log("inside " + temp.font.color);
    correct_word = "";
    console.log(term.items[cur_word].font.color);
};

const show_suggestion = (suggestions) => {
    error = true;
    $('#list').css("display", "");
    suggestions.forEach((suggestion) => {
        $("#select").append("<li>" + suggestion + "</li>");
    });

    $('#button_table').css('display', '');
    if(suggestions.length === 0){
        $("#select").append("<li>কোন পরামর্শ নেই</li>")
        for(let k = 0; k < 4; k++){
            $("#select").append("<li></li>");
        }
        return;
    }
    $("ul li").on("click", function () {
        $("ul li").removeClass('selected');
        $(this).attr('class', 'selected');
        temp_correct_word = $(this).text();

    });
    $("ul li").on("dblclick", function () {
        temp_correct_word = $(this).text();
        replace_wrong_word();

    });
}

const checker = async () => {
    console.log("Length: " + paragraphs.items.length);
    for (; cur_paragraph < paragraphs.items.length; cur_paragraph++) {
        console.log("Cur Para: " + cur_paragraph);
        console.log("Cur Word:" + cur_word);
        let paragraph = paragraphs.items[cur_paragraph];
        await console.log(paragraph);
        try {
            term = await paragraph.split(remove_strategy, true, true);
            await term.load("text,font,style");
            await word_context.sync();
            console.log("Done loading");
        } catch (e) {
            console.log("Before Word Traverse, found error");
            console.log(e);
            return;
        }

        for (; cur_word < term.items.length; cur_word++) {
            temp_correct_word = '';
            $('#spell').hide()
            $('#ignore').hide()
            if (error) {
                error = false;
                await result_for_cur_word();
                if(isRefreshClicked)return;
                continue;
            }
            current_word = term.items[cur_word];
            const context_words = await getContext(term, cur_word);
            $("#select").find("li").remove();
            $('#current_word').text(current_word.text);
            const suggestions = await get_suggestion(current_word.text, context_words);
            if (!error) continue;
            $('#spell').show();
            $('#ignore').show();
            temp.font = JSON.stringify(term.items[cur_word].font);
            temp.font = JSON.parse(temp.font);
            term.items[cur_word].font.color = "red";
            term.items[cur_word].font.underline = Word.UnderlineType.waveHeavy;
            await word_context.sync();
            show_suggestion(suggestions);
            return;

        }
        cur_word = 0;
        $('#spell').hide()
        $('#ignore').hide()
    }
};

const request_suggestion = (word, contexts) => {
    return $.ajax({
        'url': '/api/p',
        'type': 'POST',
        'dataType': 'json',
        'contentType': 'application/json',
        'data': JSON.stringify({
            word,
            contexts
        }),
        success: function (abc) {
            console.log('ok');
            // reportWordsFound('Found status ' + abc.status + ' ' + abc.suggestions);
            return abc;
        },
        error: function (abc) {
            reportWordsFound('fail ' + word + contexts);
        }
    });

}

function reportWordsFound(count) {
    var url = RedactAddin.createUrlForDialog('dialogCount.html', {count: count});
    Office.context.ui.displayDialogAsync(url,
        {height: 11, width: 12, requireHTTPS: true});
}

const get_suggestion = async (word, contexts) => {
    const res = await request_suggestion(word, contexts);
    if (res.status === 'fail') {
        error = true;
    }
    return res.suggestions;
};

const refresh = async () => {
    isRefreshClicked = true;
    await run();
    isRefreshClicked = false;
    term = '';
    cur_paragraph = 0;
    cur_word = 0;
    correct_word = false;
    current_word = '';
    temp_correct_word = '';
    temp = {};
    error = false;
    await run();
};


const replace_wrong_word = async () => {
    correct_word = temp_correct_word;
    await run();
};

async function setup() {
    await Word.run(async (context) => {
        context.document.body.clear();
        context.document.body.insertParagraph(
            "একটা কথা বলব এই বই সম্পর্কে \" আপনি যেমন করে রচনা নোট করতে চেয়েছেন বইটা ঠিক সেইভাবে তৈরি করেছি ..\" জাস্ট ডাউনলোড করে যে কোন একটি রচনার উপর চোখ বুলান তাহলেই বুঝতে পারবেন ... \"", "Start"
        );
        context.document.body.insertParagraph(
            "কিন্তু ই-বুক আপনার সারা জীবন কাজে লাগবে আপনার না লাগলেও আপনার বন্ধু বা ছোট ভাই ও বোনের ও কাজে লাগতে পারে .........",
            "End"
        );
    });
}

async function run() {
    $('#start').css('display', 'none');
    await Word.run(async (context) => {
        word_context = context;
        paragraphs = word_context.document.body.paragraphs;
        paragraphs.load();
        await word_context.sync();
        await checker();
        await word_context.sync();
    });
}

function createUrlForDialog(pageUrl, data) {
    var urlComponents = [];
    if (data) {
        for (var d in data) {
            urlComponents.push(encodeURIComponent(d) + "=" + encodeURIComponent(data[d]));
        }
    }
    return window.location.protocol + '//' + window.location.host + window.location.pathname + pageUrl + (data ? "?" : "") + urlComponents.join("&");
}

const ignore = async () => {
    await run();
};

RedactAddin.createUrlForDialog = createUrlForDialog;
RedactAddin.setup = setup;
RedactAddin.run = run;
RedactAddin.refresh = refresh;
RedactAddin.spell = replace_wrong_word;
RedactAddin.ignore = ignore;
