package Spreadsheet::SimpleExcel;

use 5.006;
use strict;
use warnings;
use Spreadsheet::WriteExcel;

require Exporter;

our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ();
our @EXPORT_OK   = ();
our @EXPORT      = qw();
our $VERSION     = '0.01';

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
  print "No Worksheet defined!" unless(defined $array[0]);
  push(@{$self->{worksheets}},[@array]);
}# end add_worksheet

sub del_worksheet{
  my ($self,$title) = @_;
  my @worksheets = grep{$_->[0] ne $title}@{$self->{worksheets}};
  $self->{worksheets} = [@worksheets];
}# end del_worksheet

sub add_row{
  my ($self,$title,$arref) = @_;
  die "Aborted: Worksheet ".$title." doesn't exist at Spreadsheet::SimpleExcel add_row()\n" unless(grep{$_->[0] eq $title}@{$self->{worksheets}});
  die "Is not an arrayref at Spreadsheet::SimpleExcel add_row()\n" unless(ref($arref) eq 'ARRAY');
  foreach my $worksheet(@{$self->{worksheets}}){
    push(@{$worksheet->[1]->{'-data'}},$arref) if($worksheet->[0] eq $title);
    last;
  }
}# end add_data

sub set_headers{
  my ($self,$title,$arref) = @_;
  die "Aborted: Worksheet ".$title." doesn't exist at Spreadsheet::SimpleExcel set_headers()\n" unless(grep{$_->[0] eq $title}@{$self->{worksheets}});
  die "Is not an arrayref at Spreadsheet::SimpleExcel set_headers()\n" unless(ref($arref) eq 'ARRAY');
  foreach my $worksheet(@{$self->{worksheets}}){
    $worksheet->[1]->{'-headers'} = $arref if($worksheet->[0] eq $title);
    last;
  }
}# end add_headers

sub add_row_at{
  my ($self,$title,$index,$arref) = @_;
  die "Aborted: Worksheet ".$title." doesn't exist at Spreadsheet::SimpleExcel add_row_at()\n" unless(grep{$_->[0] eq $title}@{$self->{worksheets}});
  die "Is not an arrayref at Spreadsheet::SimpleExcel add_row_at()\n" unless(ref($arref) eq 'ARRAY');
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      my @array = @{$worksheet->[1]->{'-data'}};
      die "Index not in Array at Spreadsheet::SimpleExcel add_row_at()\n" if($index =~ /[^\d]/ || $index > $#array);
      splice(@array,$index,0,$arref);
      $worksheet->[1]->{'-data'} = \@array;
      last;
    }
  }
}# end add_row_at

sub sort_data{
  my ($self,$title,$index,$type) = @_;
  die "Aborted: Worksheet ".$title." doesn't exist at Spreadsheet::SimpleExcel sort_data()\n" unless(grep{$_->[0] eq $title}@{$self->{worksheets}});
  foreach my $worksheet(@{$self->{worksheets}}){
    if($worksheet->[0] eq $title){
      my @array = @{$worksheet->[1]->{'-data'}};
      die "Index not in Array at Spreadsheet::SimpleExcel sort_data()\n" if($index =~ /[^\d]/ || $index > $#array);
      if(_is_numeric(\@array)){
        @array = sort{$a->[$index] <=> $b->[$index]}@array;
      }
      else{
        @array = sort{$a->[$index] cmp $b->[$index]}@array;
      }
      @array = reverse(@array) if($type eq 'DESC');
      $worksheet->[1]->{'-data'} = \@array;
      last;
    }
  }
}# end sort_data

sub _is_numeric{
  my ($arref) = @_;
  foreach(@$arref){
    return 0 if($_ =~ /[^\d\.]/);
  }
  return 1;
}# end _is_numeric

sub output{
  my ($self) = @_;

  print "Content-type: ".$self->{type}."\n\n";
  my $EXCEL = new Spreadsheet::WriteExcel(\*STDOUT);

  foreach my $worksheet(@{$self->{worksheets}}){
    my $sheet = $EXCEL->addworksheet($worksheet->[0]);
    my $col = 0;
    my $row = 0;
    foreach(@{$worksheet->[1]->{-headers}}){
      $sheet->write($row,$col,$_);
      $col++;
    }
    $row++ if(scalar(@{$worksheet->[1]->{'-headers'}}) > 0);
    foreach my $data(@{$worksheet->[1]->{-data}}){
      $col = 0;
      foreach my $value(@$data){
        $sheet->write($row,$col,$value);
        $col++;
      }
      $row++;
    }
  }
  $EXCEL->close();
}# end output


# Preloaded methods go here.

1;
__END__
# Below is stub documentation for your module. You'd better edit it!

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

  # create the spreadsheet
  $excel->output();

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

=head2 output

  $excel2->output();

prints the worksheet to the STDOUT.

=head1 DEPENDENCIES

This module requires Spreadsheet::WriteExcel

=head1 BUGS

I'm sure there are some bugs in this module. Feel free to contact me if you
experienced any problem.

=head1 SEE ALSO

Spreadsheet::WriteExcel

=head1 AUTHOR

Renee Baecker, E<lt>module@renee-baecker.deE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2004 by Renee Baecker

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.6.1 or,
at your option, any later version of Perl 5 you may have available.


=cut
