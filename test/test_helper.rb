ENV["RAILS_ENV"] = "test"

require File.expand_path("../dummy_app/config/environment.rb",  __FILE__)

migration_path = Rails.root.join('db/migrate')
if ActiveRecord.gem_version >= ::Gem::Version.new("6.0.0")
  ActiveRecord::MigrationContext.new(migration_path, ActiveRecord::SchemaMigration).migrate
elsif ActiveRecord.gem_version >= ::Gem::Version.new("5.2.0")
  ActiveRecord::MigrationContext.new(migration_path).migrate
else
  ActiveRecord::Migrator.migrate(migration_path)
end

require "rails/test_help"

Rails.backtrace_cleaner.remove_silencers!

class ActiveSupport::TestCase
  # Setup all fixtures in test/fixtures/*.yml for all tests in alphabetical order.
  fixtures :all
end

require 'minitest/reporters'
Minitest::Reporters.use!

require 'helper_methods'

### Cleanup old test spreadsheets
path = File.expand_path("../../tmp", __FILE__)
FileUtils.remove_dir(path, true)
FileUtils.mkdir(path)
