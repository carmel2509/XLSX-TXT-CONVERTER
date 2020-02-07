#!/usr/bin/perl
use strict;
use warnings;
use Spreadsheet::Read;
# force install Spreadsheet::XLSX if parse XLSX error
sub xlsxConv{
	my $Excelfilename = @_[0];
	my $Textfilename = @_[1];
	my $Wb  = ReadData($Excelfilename); # Reads workbook
	my $No_row = $Wb->[1]{maxrow}; # Number of active rows in sheet 1 (Returned as non-int)
	my $No_col = $Wb->[1]{maxcol}; # Number of active columns in sheet 1

	my $Row = 1;  # Initialise row number
	my $Col = 1; 

	my @Rows = (); # Create an empty array

	# Opening file
	open (FH, ">", $Textfilename); 

	for ( $Row = 1; $Row <= int($No_row); $Row = $Row + 1 ){  # Initialisation, condition, increase by
		for ( $Col = 1; $Col <= int($No_col); $Col = $Col + 1 ){ 
			my @Selected_cell= $Wb->[1]{cell}[$Col][$Row]; # Access cell position (col,row) from sheet 1
			push @Rows, (@Selected_cell, "\t \t");
		}
		push @Rows, "\n";
	}
	
	print FH @Rows;
	
	# Closing the file 
	close(FH);

}

