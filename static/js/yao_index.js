$(document).ready(function () {
    $("#loading").hide();

    $("#commit_upload_files").click(function () {

        $("#loading").show();

        var files = $("#files_upload").get(0).files;

        if (files.length == 0) {
            M.toast({ html: "请先选择上传文件", classes: "red" });

            $("#loading").hide();
            return;
        }

        for (var i = 0; i < files.length; i++) {
            if ((!files[i].name.toLowerCase().endsWith(".xls"))
                && (!files[i].name.toLowerCase().endsWith(".xlsx"))) {

                M.toast({ html: "请确认选择上传的是excel文件（后缀名是.xls或者.xlsx）", classes: "red" });

                $("#loading").hide();
                return;
            }
        }

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