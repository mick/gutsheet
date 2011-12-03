use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'gutsheet',
    AUTHOR              => q{YOUR NAME <youremail@example.com>},
    VERSION_FROM        => 'lib/gutsheet.pm',
    ABSTRACT            => 'YOUR APPLICATION ABSTRACT',
    ($ExtUtils::MakeMaker::VERSION >= 6.3002
      ? ('LICENSE'=> 'perl')
      : ()),
    PL_FILES            => {},
    PREREQ_PM => {
        'Test::More' => 0,
        'YAML'       => 0,
        'Dancer'     => 1.3080,
        'Plack'      => 0,
        'JSON'       => 0,
        'Spreadsheet::Read' => 0,
        'Spreadsheet::ReadSXC' => 0,
        'Spreadsheet::ParseExcel' => 0,
        'Spreadsheet::XLSX' => 0,
        'Text::CSV_XS' => 0,
        'DateTime::Format::Excel' => 0,
        'LWP::UserAgent' => 0,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'gutsheet-*' },
);
