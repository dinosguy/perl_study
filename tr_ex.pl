#!/usr/bin/perl
use strict;
use warnings;
use Encode;
use Spreadsheet::ParseExcel; # Excel 2003
use Spreadsheet::WriteExcel; 
use Getopt::Long;
use Pod::Usage;

#include libary

# Defined variable;
my $ofile_name = "" ;
my $book;
my $ver = "Dino's Version1.0";

my $rst =GetOptions (
   "v|ver|version"      => \&rver,
   "h|help"             => \my $help,
  "i=s"        => \my $in_file,
  "print_mac"          => \my $print_mac,
);

my $parser      = Spreadsheet::ParseExcel->new();
my $workbook    = $parser->parse($in_file);
my $worksheet;

my @sheets;
my @targetSheet;
my @sheetCol0;
my @sheetCol1;
my @sheetCol2;
my @sheetCol3;
my @sheetCol4;
my @sheetCol5;
my @sheetCol6;
my $Ysana = 8707239000457;


my $row_min="";
my $row_max="";
my $col_min="";
my $col_max="";
my $cnt=0;

my $cell="";
my $row=0;
my $col=0;

if ($in_file){
  print "input file: $in_file\n";
  &GetSheet();
  print " row min/max: $row_min, $row_max\n";
  print " col min/max: $col_min, $col_max\n";

  $in_file =~ m{(\w+)};
  $ofile_name ="$1_v0.xls";
  if(-e $ofile_name) {
   print "exist \n";
   system("del $ofile_name");
  }

  &MakeSheet();
  exit(1);
}
if ( !defined $workbook ) {
    die $parser->error(), ".\n";
}

if ($help){
    print "\n"x3;
    usage(1,0);
    print "help!!! \n";
    }

sub rver {
    print "\n"x3;
    print "!$: $ver\n";
    print "\n"x3;
    exit(1);
    }
if ($#ARGV <0 && !$print_mac ) {
    &cur;
    usage(0,2);
    exit (1);
}

sub GetSheet {
 for $worksheet ( $workbook->worksheets() ) {
   print " [$cnt] \t", $worksheet->get_name(), "\n";
   $sheets[$cnt] = $worksheet->get_name();
   $cnt =$cnt + 1;
 }
 print "\n##################\n\n";
 print "select sheet >";
 print "";
 my $sel = <STDIN>+ 0;
 print "$sheets[$sel] is selected\n";
 
 ($worksheet) = grep { $_->get_name() eq $sheets[$sel] } $workbook ->worksheets();
 ($row_min, $row_max) = $worksheet->row_range();
 ($col_min, $col_max) = $worksheet->col_range();

 $cnt = 0;
 $row = 0;
 $col = 0;

  for $row ($row_min .. $row_max){
    $cell = $worksheet->get_cell($row,3);
     next unless $cell;

   
    $cell = $worksheet->get_cell($row,0);
   if($cell) { $sheetCol0[$cnt] = $cell->value();}
   else      { $sheetCol0[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,1);
   if($cell) { $sheetCol1[$cnt] = $cell->value();}
   else      { $sheetCol1[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,2);
   if($cell) { $sheetCol2[$cnt] = $cell->value();}
   else      { $sheetCol2[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,3);
   if($cell) { $sheetCol3[$cnt] = $cell->value();}
   else      { $sheetCol3[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,4);
   if($cell) { $sheetCol4[$cnt] = $cell->value();}
   else      { $sheetCol4[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,5);
   if($cell) { $sheetCol5[$cnt] = $cell->value();}
   else      { $sheetCol5[$cnt] = "";            }
    $cell = $worksheet->get_cell($row,6);
   if($cell) { $sheetCol6[$cnt] = $cell->value();}
   else      { $sheetCol6[$cnt] = "";            }

   #print "$sheetCol0[$cnt]\n";
   $cnt =$cnt + 1;
 } # for loop end

} #GetSheet end

sub MakeSheet {

    print "output file name:  $ofile_name\n";

    my $Crent_date="";
    my $Nxt_date="";
    my $person = "";
    my $min_time ="";
    my $workbook_0  = Spreadsheet::WriteExcel->new($ofile_name);
    my $worksheet_0 = $workbook_0->add_worksheet("new_sheet");
    my $format = $workbook_0->add_format();

    $format->set_align("center");
    $format->set_size(10);
    $worksheet_0 -> set_column('A:A', 16);
    $worksheet_0 -> set_column('B:B', 13);
    $worksheet_0 -> set_column('D:D', 13);

    $cnt=0;
    my $cnt1 =0;
    
    foreach $row ($row_min .. $row_max){
        if($row > 0 and $row < $row_max-1){
            ($Crent_date, $Nxt_date) = ($sheetCol0[$row],$sheetCol0[$row+1]);
            $Crent_date =~ /(\d+)\w(\d+)\w(\d+)/;
            $Crent_date = "$1-$2-$3";
            #print "Current_date: $Crent_date\n";
            $Nxt_date   =~ m{(\d+)\w(\d+)\w(\d+)};
            $Nxt_date   = "$1-$2-$3";
            print "Current_date: $Crent_date, Next date: $Nxt_date\n";
        }
        $worksheet_0 -> write($row+10,0, $sheetCol0[$row]);
        $worksheet_0 -> write($row+10,1, $sheetCol1[$row]);
        $worksheet_0 -> write($row+10,3, $sheetCol3[$row]);

        $cnt1 = $cnt1 +1;
    }

    #   foreach $row (1 .. $row_max){
    #       ($Crent_date, $Nxt_date) = ($sheetCol0[$row],$sheetCol0[$row+1]);
    #       #print " $Crent_date, $Nxt_date\n";
    #       $Crent_date =~ m{(\d+)\w(\d+)\w(\d+)};
    #       $Crent_date = "$1-$2-$3";
    #
    #       #exit(1);
    #       #$Nxt_date = &Nxt_d($Nxt_date);
    #       $Nxt_date   =~ m{(\d+)\w(\d+)\w(\d+)};
    #       $Nxt_date   = "$1-$2-$3";
    #       if ($row eq 1){
    #           $worksheet_0 -> write($cnt1+10,0, $Crent_date);
    #           $worksheet_0 -> write($cnt1+10,1, $sheetCol1[$row]);
    #           $worksheet_0 -> write($cnt1+10,3, $sheetCol3[$row]);
    #       }
    #       if($Crent_date ne $Nxt_date) {
    #           $worksheet_0 -> write($cnt1+10,0, $Crent_date);
    #           $worksheet_0 -> write($cnt1+10,1, $sheetCol1[$row]);
    #           $worksheet_0 -> write($cnt1+10,3, $sheetCol3[$row]);
    #
    #           #print "different date: $Crent_date,$Nxt_date\n";
    #       }
    #
    #       # Ysana
    #       if ($sheetCol2[$row] eq "$Ysana"){
    #
    #           $min_time = $sheetCol1[$row];
    #           #if ($min_time < $sheetCol
    #
    #           # $worksheet_0 -> write($cnt1+10,2, $sheetCol2[$row]);
    #           $cnt1= $cnt1 +1;
    #       }
    #       elsif($row<1){
    #          print" row: $row\n";
    #       }
    #       $cnt = $cnt +1;
    #   }
 print "cnt1 value: $cnt1\n";

}

#sub Nxt_d{
#    @_ =~ m{(\d+)\w(\d+)\w(\d+)};
#    my $Nxt_temp = "$1-$2-$3";
#
#    return $Nxt_temp;
#}
#$worksheet_0->write(0,0, 'Hi Excel!');
#$worksheet_0->write(0,1, 'Hi Excel!');



sub usage {
    my ($verbose, $exitval) =@_;
    pod2usage ( -verbose => $verbose, -exitval => $exitval);
    exit(1);
}


sub cur{
  warn 
   "\n"x2;
  warn 
   "\n"x2;
}

__DATA__

=pod

=head1 NAME

=head1 SYNOPSIS
 tr_ex.pl : [option -i, -h, -v]

=cut

