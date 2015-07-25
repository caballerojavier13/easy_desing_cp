require 'test_helper'

class VisualReDesignTestControllerTest < ActionController::TestCase
  test "should get index" do
    get :index
    assert_response :success
  end

  test "should get download" do
    get :download
    assert_response :success
  end

end
