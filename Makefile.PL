use 5.006001;
use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'Spreadsheet::SimpleExcel',
    VERSION_FROM      => 'lib/Spreadsheet/SimpleExcel.pm', # finds $VERSION
    PREREQ_PM         => {
            'Spreadsheet::WriteExcel'    => '0', # or a minimum workable version
            'IO::Scalar'                 => '0',
            'IO::File'                   => '1.10',
            'XML::Writer'                => '0.600'
    },
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      (ABSTRACT_FROM  => 'lib/Spreadsheet/SimpleExcel.pm', # retrieve abstract from module
       AUTHOR         => 'Renee Baecker <module@renee-baecker.de>') : ()),
);
