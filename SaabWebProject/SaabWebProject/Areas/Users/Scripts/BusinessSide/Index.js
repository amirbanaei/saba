///Show Partial View Create Business
function CreateBusinessSide(url) {
    $(".modal-body").load(url);
}

/// Show Partial View Edit Business
function EditBusinessSide(url, input) {
    var id = input.children[0].value;
    $.ajax({
        url: url,
        type: "get",
        data: {id: id},
        dataType: "html",
        success: function (res) {
            $(".modal-body").html(res);
        }
    });
}

function AccessLevel(url, input) {
    var id = input.children[0].value;
    $("#big_modal_body").load(url + '?Id=' + id);
}

/// DeActive Or Active a Business
function DeActiveOrActive(url, input) {

    var state = "غیرفعال";

    if ($(input).hasClass("false")) {
        state = "فعال";
    }
    $(input).html(state);
    var up = $(input).attr('id').split('_')[1];
    console.log(up);
    Swal.fire({
        title: 'آیا مطمئن هستید؟',
        text: "این سمت در کل برنامه " + state + " می شود",
        icon: 'question',
        showCancelButton: true,
        background: 'rgb(29,40,52)',
        color: 'white',
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: 'بله, ' + state + ' کن!',
        cancelButtonText: '' + state + ' نکن'
    }).then((result) => {
        if (result.isConfirmed) {
            input.setAttribute('title', state);
            $.ajax({
                url: url,
                type: "get",
                data: {id: up},
                dataType: "html",
                success: function (res) {
                    if (res == "False") {
                        Swal.fire({
                            icon: 'error',
                            background: 'rgb(29,40,52)',
                            color: 'white',
                            confirmButtonText: 'باشه',
                            title: 'خطا',
                            text: 'برنامه با خطا مواجه شد',
                        });
                        return false;
                    }
                    Swal.fire({
                        icon: 'success',
                        background: 'rgb(29,40,52)',
                        color: 'white',
                        confirmButtonText: 'باشه',
                        title: '' + state + ' شد ',
                        text: 'سمت مورد نظر ' + state + ' شد',
                    }).then(() => {
                        location.reload();
                    });


                    if ($(input).hasClass("true")) {
                        $(input).removeClass("true");
                        $(input).addClass("false");

                    } else {
                        $(input).removeClass("false");
                        $(input).addClass("true");
                    }
                }
            });
        }
    })
}