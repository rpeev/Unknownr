require_relative '../lib/unknownr'
require 'minitest'
require 'minitest/autorun'

class UnknownrTest < Minitest::Test
  def test_version
    refute_nil Unknownr::VERSION
  end
end
