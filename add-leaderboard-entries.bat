SET PATH=C:\Ruby193\bin;%PATH
pushd %~dp0
bundle exec ruby add-leaderboard-entries.rb
popd
