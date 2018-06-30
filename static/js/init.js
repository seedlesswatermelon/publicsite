// (
//   function ($) {
//     $(function () {

//       $('.sidenav').sidenav();

//     }); // end of document ready
//   }
// )(jQuery); // end of jQuery name space

$(document).ready(function () {
    $("#loading").hide();

    $("#commit_upload_files").click(function () {

        $("#loading").show();

        var files = $("#files_upload").get(0).files;
        var formData = new FormData();
        for (var i = 0; i < files.length; i++) {
            formData.append("file", files[i]);
        }

        $.ajax(
            {
                url: "/upload",
                type: "POST",
                async: false,

                data: formData,

                cache: false,
                contentType: false,
                processData: false

            }).success(function (html) {
            $("body").html(html);
        })
    });
});