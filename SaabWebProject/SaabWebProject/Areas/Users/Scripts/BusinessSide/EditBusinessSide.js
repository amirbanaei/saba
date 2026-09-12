
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


