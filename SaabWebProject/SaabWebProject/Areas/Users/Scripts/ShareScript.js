
/// تابع ولید سنجی تابع شماره همراه
$("#usr_NationalCode").keyup(function () {

    var PhoneNumber = document.getElementById("usr_NationalCode");
    var PhoneNumber_Length = PhoneNumber.value.length;
    if (PhoneNumber_Length > 10 || PhoneNumber_Length < 10) {
        if (PhoneNumber.classList.contains("is-valid")) {
            PhoneNumber.classList.remove("is-valid");
            PhoneNumber.classList.add("is-invalid");

        }
        else {
            PhoneNumber.classList.add("is-invalid");

        }
        return false;
    }
    else {
        if (PhoneNumber.classList.contains("is-invalid")) {
            PhoneNumber.classList.remove("is-invalid");
            PhoneNumber.classList.add("is-valid");
        }
        else {
            PhoneNumber.classList.add("is-valid");
        }
    }

});