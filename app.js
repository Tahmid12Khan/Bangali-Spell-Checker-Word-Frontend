(function() {
    "use strict";
    console.log('here');

    // The initialize function is run each time a page of the add-in is loaded into the task pane.
    Office.initialize = function(reason) {
        $(document).ready(function() {

            $.ajax({
                url: '/api?what=abc2',
                success: function (abc) {
                    console.log('ok')
                }
            });

            // Use this to check whether the new API is supported in the Word client.......
            // The createDocument method call is in the 1.3 requirement set. 
            // The 1.3 requirement set check is not implemented in preview. 
            // The 1.2 API requirement set check is the minimum requirement check in Word. 
            // Update this to target the correct set after 1.3 is generally available.
            if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {
                $('#setup').click(RedactAddin.setup)
                $('#start').click(RedactAddin.run);
                $('#spell').click(RedactAddin.spell);
                $('#refresh').click(RedactAddin.refresh);
                $('#ignore').click(RedactAddin.ignore);
            }
            else {                
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires the latest version of Word!!');
            }
        });
    };    
})();
