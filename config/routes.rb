FacilitadorDeDiseOWeb::Application.routes.draw do

  get 'visual_re_design_test/index'

  post 'visual_re_design_test/download'

  post 'design_cp/download' => 'design_cp#download'

  post 'design_cp/upload' => 'design_cp#upload'

  post 'design_cp/upload_confirmation' => 'design_cp#upload_confirmation'

  get 'design_cp' => 'design_cp#index'

  get '' => 'home#index'

  root 'home#index'

end
