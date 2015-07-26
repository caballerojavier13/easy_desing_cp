//= require jquery
//= require jquery_ujs
//= require turbolinks
//= require bootstrap.min
//= require select2.min


var miniLoading = false;

$(function() {
    $(".select2").select2();
});

$(document).on('change', '.btn-file', function() {
    var input = $(this),
        numFiles = input.get(0).files ? input.get(0).files.length : 1,
        label = input.val().replace(/\\/g, '/').replace(/.*\//, '');
    input.trigger('fileselect', [numFiles, label]);
});

$(document).ready( function() {
    $('.btn-file').on('fileselect', function(event, numFiles, label) {

        var input = $(this).parents('.input-group').find(':text'),
            log = numFiles > 1 ? numFiles + ' files selected' : label;

        if( input.length ) {
            input.val(log);
        } else {
            if( log ) alert(log);
        }
    });

    $('.upload').on('click', function(){
        showMiniLoading();
    });
});


function showMiniLoading() {
    miniLoading = true;
    $("body").append('<svg class="spinner" id="mini_loading" width="65px" height="65px" viewBox="0 0 66 66" xmlns="http://www.w3.org/2000/svg"><circle class="path" fill="none" stroke-width="3" stroke-linecap="round" cx="33" cy="33" r="30"></circle></svg>');
}

function hideMiniLoading() {
    if(miniLoading) {
        setTimeout(
            function() {
                $("#mini_loading").remove();
            }, 2000);
    }
}