Gem::Specification.new do |s|
	s.name="xlutowin"
	s.version="0.0.1"
	s.authors=["amir kouretchian"]
	s.date=%q{2013-11-20}
	s.description="quickly read excel's \"unicode\" files and transcode their contents to windows-1252 (by default.)"
	s.summary=s.description
	s.files=%w(README.md).concat(Dir.glob("**/*.rb"))
end
