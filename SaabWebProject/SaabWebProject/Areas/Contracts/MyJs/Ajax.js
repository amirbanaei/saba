function sendAjax(url, data) {

    $.ajax({
        beforeSend:function (){
            document.getElementById("loading-message").style.display = "block";  
        },
        url: url,
        data: { Filters: data },
        type: "Post",
        dataType: "html",
        success: function (resp) {
            document.getElementById("loading-message").style.display = "none";
            return resp;
        },
        error: function (resp) {
            return resp;
        }
    });

}
