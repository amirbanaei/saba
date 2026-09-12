$("#Other").click(function () {
    if ($("#Other_Arrow").hasClass("fa-angle-up")) {
        $("#Other_Arrow").removeClass("fa-angle-up");
        $("#Other_Arrow").addClass("fa-angle-down");
        $("#Other_div").css("display", "block");
    }
    else {
        $("#Other_Arrow").removeClass("fa-angle-down");
        $("#Other_Arrow").addClass("fa-angle-up");
        $("#Other_div").css("display", "none");
    }

});

$("#btnSave").click(function () {
    var up_name = $("input[name='up_name']");
    if (up_name.val() == null || up_name.val() == "") {
        up_name.addClass("is-invalid");
    }
    else {
        up_name.remove("is-invalid");
        document.getElementById("loading-message").style.display = "block";
        $('#CreateForm').submit();
    }
});