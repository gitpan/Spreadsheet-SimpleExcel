package Spreadsheet::SimpleExcel;

use 5.006;
use strict;
use warnings;
use Spreadsheet::WriteExcel;
use IO::Scalar;

require Exporter;

our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ();
our @EXPORT_OK   = ();
our @EXPORT      = qw();
our $VERSION     = '0.4';
our $errstr      = '';

sub new{
  my ($class,%opts) = @_;
  my $self = {};
  $self->{worksheets} = $opts{-worksheets} || [];
  $self->{type}       = 'application/vnd.ms-excel';
  bless($self,$class);
  return $self;
}# end new

sub add_worksheet{
  my ($self,@array) = @_;
  my ($package,$filename,$line) = caller();
  unless(defined $array[0]){
    $errstr = qq~No worksheet defined at Spreadsheet::SimpleExcel add_worksheet() from
       $filename line $line\n~;
    $array[0] ||= 'unknown';
    return undef;
  }
  if(grep{$_->[0] eq $array[0]}@{$self->{worksheets}}){
    $errstr = qq~Duplicate worksheet-title at Spreadsheet::SimpleExcel add_worksheet() from
       $filename line $line\n~;
    return undef;
  }
  push(@{$self->{worksheets}},[@array]);
  return 1;
}# end add_worksheet

sub del_worksheet{
  my ($self,$title) = @_;
  my ($package,$filename,$line) = caller();
  unless(defined $title){
    $errstr = qq~No worksheet-title defined at Spreadsheet::SimpleExcel del_worksheet() from
        $filename line $line\n~;
    return undef;
  }
  my @worksheets = grep{$_->[0] ne $title}@{$self->{worksheets}};
  $self->{worksheets} = [@worksheets];
}# end del_worksheet

sub add_row{
  my ($self,$title,$arref) = @_;
  my ($package,$filename,$line) = caller();
  $title ||= 'unknown';
  unless(grep{$_->[0] eq $title}@{$self->{worksheets}}){
    $errstr = qq~Worksheet $title does not exist at Spreadsheet::SimpleExcel add_row() from
         $filename line $line\n~;
    return undef;
  }
  unless(ref($arref) eq 'ARRAY'){
    $errstr = qq~Is not an arrayref at Spreadsheet::SimpleExcel add_row() from
         $filename line $line\n~;
    return undef;
  }
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      push(@{$worksheet->[1]->{'-data'}},$arref);
      last;
    }
  }
  return 1;
}# end add_data

sub set_headers{
  my ($self,$title,$arref) = @_;
  $title ||= 'unknown';
  my ($package,$filename,$line) = caller();
  unless(grep{$_->[0] eq $title}@{$self->{worksheets}}){
    $errstr = qq~Worksheet $title does not exist at Spreadsheet::SimpleExcel set_headers() from
         $filename line $line\n~;
    return undef;
  }
  unless(ref($arref) eq 'ARRAY'){
    $errstr = qq~Is not an arrayref at Spreadsheet::SimpleExcel add_row() from
         $filename line $line\n~;
    return undef;
  }
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      $worksheet->[1]->{'-headers'} = $arref;
      last;
    }
  }
  return 1;
}# end add_headers

sub add_row_at{
  my ($self,$title,$index,$arref) = @_;
  my ($package,$filename,$line) = caller();
  $title ||= 'unknown';
  unless(grep{$_->[0] eq $title}@{$self->{worksheets}}){
    $errstr = qq~Worksheet $title does not exist at Spreadsheet::SimpleExcel add_row_at() from
         $filename line $line\n~;
    return undef;
  }
  unless(ref($arref) eq 'ARRAY'){
    $errstr = qq~Is not an arrayref at Spreadsheet::SimpleExcel add_row() from
         $filename line $line\n~;
    return undef;
  }
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      my @array = @{$worksheet->[1]->{'-data'}};
      if($index =~ /[^\d]/ || $index > $#array){
        $errstr = qq~Index not in Array at Spreadsheet::SimpleExcel add_row_at() from
         $filename line $line\n~;
        return undef;
      }
      splice(@array,$index,0,$arref);
      $worksheet->[1]->{'-data'} = \@array;
      last;
    }
  }
  return 1;
}# end add_row_at

sub sort_data{
  my ($self,$title,$index,$type) = @_;
  my ($package,$filename,$line) = caller();
  $title ||= 'unknown';
  $type  ||= 'ASC';
  unless(grep{$_->[0] eq $title}@{$self->{worksheets}}){
    $errstr = qq~Worksheet $title does not exist at Spreadsheet::SimpleExcel sort_data() from
          $filename line $line\n~;
    return undef;
  }
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      my @array = @{$worksheet->[1]->{'-data'}};
      if(not defined $index || $index =~ /[^\d]/ || $index > $#array){
        $errstr = qq~Index not in Array at Spreadsheet::SimpleExcel sort_data() from
          $filename line $line\n~;
        return undef;
      }
      if(_is_numeric(\@array)){
        @array = sort{$a->[$index] <=> $b->[$index]}@array;
      }
      else{
        @array = sort{$a->[$index] cmp $b->[$index]}@array;
      }
      @array = reverse(@array) if($type && $type eq 'DESC');
      $worksheet->[1]->{'-data'} = \@array;
      last;
    }
  }
  return 1;
}# end sort_data

sub errstr{
  return $errstr;
}# end errstr

sub sort_worksheets{
  my ($self,$type) = @_;
  $type ||= 'ASC';
  my @title_array = map{$_->[0]}@{$self->{worksheets}};
  if(_is_numeric(\@title_array)){
    @{$self->{worksheets}} = sort{$a->[0] <=> $b->[0]}@{$self->{worksheets}};
  }
  else{
    @{$self->{worksheets}} = sort{$a->[0] cmp $b->[0]}@{$self->{worksheets}};
  }
  @{$self->{worksheets}} = reverse(@{$self->{worksheets}}) if($type && $type eq 'DESC');
  return @{$self->{worksheets}};
}# end sort_worksheets

sub _is_numeric{
  my ($arref) = @_;
  foreach(@$arref){
    return 0 if($_ =~ /[^\d\.]/);
  }
  return 1;
}# end _is_numeric

sub output{
  my ($self,$lines) = @_;
  my ($package,$filename,$line) = caller();
  $lines ||= 32000;
  $lines =~ s/\D//g;
  my $excel = $self->_make_excel($lines);
  unless(defined $excel){
    $errstr = qq~Could not create Spreadsheet at Spreadsheet::SimpleExcel output() from
         $filename line $line\n~;
    return undef;
  }
  print "Content-type: ".$self->{type}."\n\n",
        $excel;
}# end output

sub output_as_string{
  my ($self,$lines) = @_;
  my ($package,$filename,$line) = caller();
  $lines ||= 32000;
  $lines =~ s/\D//g;
  my $excel = $self->_make_excel($lines);
  unless(defined $excel){
    $errstr = qq~Could not create Spreadsheet at Spreadsheet::SimpleExcel output_to_file() from
        $filename line $line\n~;
    return undef;
  }
  return $excel;
}# end output_as_string

sub output_to_file{
  my ($self,$filename,$lines) = @_;
  my ($package,$file,$line) = caller();
  $lines ||= 32000;
  $lines =~ s/\D//g;
  unless($filename){
    $errstr = qq~No filename specified at Spreadsheet::SimpleExcel output_to_file() from
        $file line $line\n~;
    return undef;
  }
  $filename =~ s/[^A-Za-z0-9_\.\/]//g; #/
  my $excel = $self->_make_excel($lines);
  unless(defined $excel){
    $errstr = qq~Could not create $filename at Spreadsheet::SimpleExcel output_to_file() from
        $file line $line\n~;
    return undef;
  }
  open(EXCEL,">$filename") or die $!;
  print EXCEL $excel;
  close EXCEL;
  return 1;
}# end output_to_file

sub _make_excel{
  my ($self,$nr_of_lines) = @_;
  my ($package,$filename,$line) = caller();
  my $c_lines = $nr_of_lines || 32000;
  unless(scalar(@{$self->{worksheets}}) >= 1){
    $errstr = qq~No worksheets in Spreadsheet~;
    return undef;
  }
  my $output;
  tie(*XLS,'IO::Scalar',\$output);
  my $excel;
  unless($excel = new Spreadsheet::WriteExcel(\*XLS)){
    $errstr = qq~Could not create spreadsheet object ($!) from
        $filename line $line~;
    return undef;
  }
  else{
    my @titles = map{$_->[0]}@{$self->{worksheets}};
    foreach my $worksheet(@{$self->{worksheets}}){
      my $sheet = $excel->addworksheet($worksheet->[0]);
      my $col  = 0;
      my $row  = 0;
      my $page = 2;
      _header2sheet($sheet,$worksheet->[1]->{-headers});
      $row++ if(exists $worksheet->[1]->{'-headers'} && scalar(@{$worksheet->[1]->{'-headers'}}) > 0);
      foreach my $data(@{$worksheet->[1]->{-data}}){
        $col = 0;
        if($row >= $c_lines){
          my $title = $worksheet->[0].'_p'.$page;
          while(grep{$_ eq $title}@titles){
            $page++;
            $title = $worksheet->[0].'_p'.$page;
          }
          push(@titles,$title);
          $sheet = $excel->addworksheet($title);
          $row = 0;
          if(scalar(@{$worksheet->[1]->{'-headers'}}) > 0){
            $row = 1;
            _header2sheet($sheet,$worksheet->[1]->{-headers});
          }
        }
        foreach my $value(@$data){
          $sheet->write($row,$col,$value);
          $col++;
        }
        $row++;
      }
    }
    $excel->close();
  }
  return $output;
}# end _make_excel

sub _header2sheet{
  my ($sheet,$arref) = @_;
  my $col = 0;
  foreach(@$arref){
    $sheet->write(0,$col,$_);
    $col++;
  }
}# end _header2sheet

sub sheets{
  my ($self) = @_;
  my @titles = map{$_->[0]}@{$self->{worksheets}};
  return wantarray ? @titles : \@titles;
}# end sheets

1;
__END__

=head1 NAME

Spreadsheet::SimpleExcel - Perl extension for creating excel-files quickly

=head1 SYNOPSIS

  use Spreadsheet::SimpleExcel;

  binmode(\*STDOUT);
  # data for spreadsheet
  my @header = qw(Header1 Header2);
  my @data   = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  # create a new instance
  my $excel = Spreadsheet::SimpleExcel->new();

  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_worksheet('Second Worksheet',{-data => \@data});
  $excel->add_worksheet('Test');

  # add a row into the middle
  $excel->add_row_at('Name of Worksheet',1,[qw/new row/]);

  # sort data of worksheet - ASC or DESC
  $excel->sort_data('Name of Worksheet',0,'DESC');

  # remove a worksheet
  $excel->del_worksheet('Test');

  # sort worksheets
  $excel->sort_worksheets('DESC');

  # create the spreadsheet
  $excel->output();

  # print sheet-names
  print join(", ",$excel->sheets()),"\n";

  # get the result as a string
  my $spreadsheet = $excel->output_as_string();

  # print result into a file and handle error
  $excel->output_to_file("my_excel.xls") or die $excel->errstr();
  $excel->output_to_file("my_excel2.xls",45000) or die $excel->errstr();

  ## or

  # data
  my @data2  = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  my $worksheet = ['NAME',{-data => \@data2}];
  # create a new instance
  my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

  # add headers to 'NAME'
  $excel2->set_headers('NAME',[qw/this is a test/]);
  # append data to 'NAME'
  $excel2->add_row('NAME',[qw/new row/]);

  $excel2->output();

=head1 DESCRIPTION

Spreadsheet::SimpleExcel simplifies the creation of excel-files in the web. It does not
provide any access to cell-formats yet. This is just a raw version that will be
extended within the next few weeks.

=head1 METHODS

=head2 new

  # create a new instance
  my $excel = Spreadsheet::SimpleExcel->new();

  # or

  my $worksheet = ['NAME',{-data => ['This','is','an','Test']}];
  my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

  # to create a file
  my $filename = 'test.xls';
  my $excel = Spreadsheet::SimpleExcel->new(-filename => $filename);

=head2 add_worksheet

  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_worksheet('Second Worksheet',{-data => \@data});
  $excel->add_worksheet('Test');

The first parameter of this method is the name of the worksheet and the second one is
a hash with (optional) information about the headlines and the data.
No duplicate worksheets allowed.

=head2 del_worksheet

  # remove a worksheet
  $excel->del_worksheet('Test');

Deletes all worksheets named like the first parameter

=head2 add_row

  # append data to 'NAME'
  $excel->add_row('NAME',[qw/new row/]);

Adds a new row to the worksheet named 'NAME'

=head2 add_row_at

  # add a row into the middle
  $excel->add_row_at('Name of Worksheet',1,[qw/new row/]);

This method inserts a row into the existing data

=head2 sort_data

  # sort data of worksheet - ASC or DESC
  $excel->sort_data('Name of Worksheet',0,'DESC');

sort_data sorts the rows

=head2 set_headers

  # add headers to 'NAME'
  $excel->set_headers('NAME',[qw/this is a test/]);

set the headers for the worksheet named 'NAME'

=head2 errstr

returns error message.

=head2 sort_worksheets

  # sort worksheets
  $excel->sort_worksheets('DESC');

sorts the worksheets in DESCending or ASCending order.

=head2 output

  $excel2->output();

prints the worksheet to the STDOUT and prints the Mime-type 'application/vnd.ms-excel'.

=head2 output_as_string

  # get the result as a string
  my $spreadsheet = $excel->output_as_string();

returns a string that contains the data in excel-format

=head2 output_to_file

  # print result into a file [output_to_file(<filename>,<lines>)]
  $excel->output_to_file("my_excel.xls");
  $excel->output_to_file("my_excel2.xls",45000) or die $excel->errstr();

prints the data into a file. There is a limitation in allowed characters for the filename:
A-Za-z0-9/._
Other characters will be deleted.
The data will be printed into more worksheets, if the number of rows is greater than <lines> (default 32000).

=head2 sheets

  $ref = $excel->sheets();
  @names = $excel->sheets();

In Listcontext this subroutines returns a list of the names of sheets that are in $excel, in
scalar context it returns a reference on an Array.

=head1 EXAMPLES

=head2 PRINT ON STDOUT

  #! /usr/bin/perl

  use strict;
  use warnings;
  use Spreadsheet::SimpleExcel;

  binmode(\*STDOUT);
  # data for spreadsheet
  my @header = qw(Header1 Header2);
  my @data   = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  # create a new instance
  my $excel = Spreadsheet::SimpleExcel->new();

  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_worksheet('Second Worksheet',{-data => \@data});
  $excel->add_worksheet('Test');

  # add a row into the middle
  $excel->add_row_at('Name of Worksheet',1,[qw/new row/]);

  # sort data of worksheet - ASC or DESC
  $excel->sort_data('Name of Worksheet',0,'DESC');

  # remove a worksheet
  $excel->del_worksheet('Test');

  # create the spreadsheet
  $excel->output();

=head2 RECEIVE DATA AS A SCALAR

  #!/usr/bin/perl

  use strict;
  use warnings;
  use Spreadsheet::SimpleExcel;

  # data
  my @data2  = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  my $worksheet = ['NAME',{-data => \@data2}];
  # create a new instance
  my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

  # add headers to 'NAME'
  $excel2->set_headers('NAME',[qw/this is a test/]);
  # append data to 'NAME'
  $excel2->add_row('NAME',[qw/new row/]);

  # receive as string
  my $string = $excel2->output_as_string();

=head2 PRINT INTO FILE

  #! /usr/bin/perl

  use strict;
  use warnings;
  use Spreadsheet::SimpleExcel;

  # data
  my @data2  = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  my $worksheet = ['NAME',{-data => \@data2}];
  # create a new instance
  my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

  # add headers to 'NAME'
  $excel2->set_headers('NAME',[qw/this is a test/]);
  # append data to 'NAME'
  $excel2->add_row('NAME',[qw/new row/]);

  # print into file
  $excel2->output_to_file("my_excel.xls");

=head2 PRINT INTO FILE (break worksheets)

  #! /usr/bin/perl

  use strict;
  use warnings;
  use Spreadsheet::SimpleExcel;

  # create a new instance
  my $excel    = Spreadsheet::SimpleExcel->new();

  my @header = qw(Header1 Header2);
  my @data   = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);
  for(0..70000){
    push(@data,[qw/1 2 4 6 8/]);
  }
  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_row('Name of Worksheet',[qw/1 2 3 4 5/]);

  # print into file
  $excel->output_to_file("my_excel.xls",10000);

=head1 DEPENDENCIES

This module requires Spreadsheet::WriteExcel and IO::Scalar

=head1 BUGS and COMMENTS

Feel free to contact me and send me bugreports or comments on this module.

=head1 SEE ALSO

Spreadsheet::WriteExcel

IO::Scalar

=head1 AUTHOR

Renee Baecker, E<lt>module@renee-baecker.deE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2004 by Renee Baecker

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.6.1 or,
at your option, any later version of Perl 5 you may have available.


=cut
