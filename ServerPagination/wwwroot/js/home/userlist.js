/********************** start pagination listing **********************/
let searchQuery = '';
let PageNumber = 1;
$(document).ready(function () {
    UserList(PageNumber);
    $('#demo-input-search2').on('input', function () {
        searchQuery = $(this).val();
        UserList(1);
    });
});
function UserList(pageNumber) {
    PageNumber = pageNumber;
    ShowLoader();
    $.ajax({
        type: 'get',
        url: '/Home/LoadUserList',
        data: { pageNumber: pageNumber, searchQuery: searchQuery },
        success: function (data) {
            HideLoader();
            $('#UserListPartial').html(data);
        },
        error: function (er) {
            console.log(er);
        },
    });
}
/********************** end pagination listing **********************/

/********************** start add modal  **********************/
function OpenAddUserForm() {
    $.ajax({
        type: 'get',
        url: '/Home/AddUser',
        data: null,
        success: function (data) {
            $("#addUserBody").html(data);
        },
        error: function (error) {
            ErrorRedirection(error);
        },
    });
}
function AddUser() {
    let formData = $('#AddUserForm').serialize();
    $.ajax({
        type: 'post',
        url: '/Home/AddUser',
        data: formData,
        success: function (data) {
            if (data == true) {
                $('#addclose').click();
                $("#AddUserForm")[0].reset();
                UserList(PageNumber);
            }
            else {
                $('#addUserBody').html(data);
            }
        },
        error: function (er) {
            console.log(er);
        },
    });
}
/********************** end add modal **********************/

/********************** start view data **********************/
function ViewUser(UserId) {
    $.ajax({
        type: 'get',
        url: '/Home/ViewUser',
        data: { UserId: UserId },
        success: function (data) {
            $('#viewUserBody').html(data);
        },
        error: function (er) {
            console.log(er);
        },
    });
}
/********************** end view data **********************/

/********************** start edit modal **********************/
function GetUser(UserId) {
    $.ajax({
        type: 'get',
        url: '/Home/GetUser',
        data: { UserId: UserId },
        success: function (data) {
            $('#editUserBody').html(data);
        },
        error: function (er) {
            console.log(er);
        },
    });
}
function EditUser() {
    let formData = $('#EditUserForm').serialize();
    $.ajax({
        type: 'put',
        url: '/Home/EditUser',
        data: formData,
        success: function (data) {
            if (data == true) {
                $('#editclose').click();
                $("#EditUserForm")[0].reset();
                UserList(PageNumber);
            }
            else {
                $('#editUserBody').html(data);
            }
        },
        error: function (er) {
            console.log(er);
        },
    });
}
/********************** End edit modal **********************/

/********************** start delete data **********************/
function DeleteUser(UserId) {
    swal({
        title: "Are you sure?",
        text: "Once deleted, you will not be able to recover this Data!",
        icon: "warning",
        buttons: true,
        dangerMode: true,
    })
        .then((willDelete) => {
            if (willDelete) {
                $.ajax({
                    type: 'delete',
                    url: '/Home/DeleteUser',
                    data: { UserId: UserId },
                    success: function (data) {
                        swal("Poof! Your Data has been deleted!", {
                            icon: "success",
                        });
                        UserList(PageNumber);
                    },
                    error: function (er) {
                        console.log(er);
                    },
                });
            } else {
                swal("Your Recode is safe!");
            }
        });
}
/********************** end delete data **********************/

/********************** start active manage data **********************/
function ActiveManage(UserId) {
    $.ajax({
        type: 'get',
        url: '/Home/ActiveManage',
        data: { UserId: UserId },
        success: function (data) {
            UserList(PageNumber);
        },
        error: function (er) {
            console.log(er);
        },
    });
}
/********************** end delete data **********************/