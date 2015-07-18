$(function(){
    $("#cu").on('change',function(){
        $("#nombre_cu").val($("#cu option:selected").text());
    });
});

