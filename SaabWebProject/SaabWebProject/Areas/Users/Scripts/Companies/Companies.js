/// Toast SweetAlert
const Toast = Swal.mixin({
    toast: true,
    showConfirmButton: false,
    timer: 3000,
    timerProgressBar: true,
    didOpen: (toast) => {
        toast.addEventListener('mouseenter', Swal.stopTimer)
        toast.addEventListener('mouseleave', Swal.resumeTimer)
    }
});

/// Button Disable Company
function Disable(url, status) {
    let message = "غیرفعال";
    if (status) {
        message = "فعال";
    }
    var title = "کاربر حقوقی با موفقیت " + message  + " شد "+ " " ;
    $.ajax({
        url: url,
        type: "Get",
        datatype: "html",
        success: function (res) {
            if (res === "True") {
                Toast.fire({
                    icon: 'success',
                    background: '#35434A',
                    color: 'white',
                    title: title
                });
                $.ajax({
                    beforeSend:function (){
                        document.getElementById("loading-message").style.display = "block";  
                    },
                    url:"/Company/_PartialListComanies",
                    type: "get",
                    dataType: "html",
                    success: function (res) {
                        document.getElementById("loading-message").style.display = "none";
                        $("#myTable").html(res);
                    }
                });
            }
        },
        error: function (res) {
            document.getElementById("loading-message").style.display = "none";
            Toast.fire({
                icon: 'warning',
                background: '#35434A',
                color: 'white',
                title: 'عملیات با خطا مواجه شد لطفا بعدا تلاش فرمائید'
            });
            console.log(res);
        }
    });
    
}

/// Button Edit Company
function Edit(url) {
    $.ajax({
        url: url,
        type: "Get",
        datatype: "html",
        success: function (res) {
            $("#big_modal_body").html(res);
        }
    });
}
