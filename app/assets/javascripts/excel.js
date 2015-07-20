$(function(){
    $("#cu").on('change',function(){
        $("#nombre_cu").val($("#cu option:selected").text());
    });

    $("#select_precondicion").on('change',function(){
        var val = $(this).val();
        if(val < 1){
            $('#accordion_precondicion').hide();
            $('#precondicion_all').show();
        }else{
            $('#accordion_precondicion').show();
            $('#precondicion_all').hide();
        }
    });

    $("#select_pasos").on('change',function(){
        var val = $(this).val();
        if(val < 1){
            $('#accordion_pasos').hide();
            $('#pasos_all').show();
        }else{
            $('#accordion_pasos').show();
            $('#pasos_all').hide();
        }
    });

});

