package ParseSheet;
use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::WriteExcel;


our $workbook;
our $worksheet;
our $writebook;
our $writesheet;
our $sheetname;
our $row_min;
our $row_max;
our $col_min;
our $col_max;

sub hashValueAscendingNum { $a <=> $b; }

sub selectFile {
	my %fileHash;
	opendir (DIR, ".");
	my @files = grep(/\.xls$/ || /\.csv$/, readdir(DIR));
	closedir(DIR);
	my $i = 0;
	for (@files) { 
		$i++;
		$fileHash{$i} = $_;
	}
	print map { "\t$_ : $fileHash{$_}\n"} sort hashValueAscendingNum( keys(%fileHash) );
	
	return \%fileHash;	
}
sub openFile {
	my $parser   = Spreadsheet::ParseExcel->new();
	$workbook = $parser->parse( $_[0] );
	if ( !defined $workbook ) { die $parser->error(), ".\n"; }

	return $workbook;
}

sub openRefSheet {
	
	my $sheetCnt;
	our $sheetname;
	
	if ( $_[0]->worksheet_count() == 1 ) { 
		$worksheet = $_[0]->worksheet(0);
	} elsif ( $_[0]->worksheet_count() == 0 ) {
		print "There is no sheet in ",$_[0]->get_filename(),"\n";
		return 0;
	} else {
		for my $sheet ( $_[0]->worksheets() ) { 
			print "\t",$sheetCnt++,") \"",$sheet->get_name(),"\"\n"; 
		}
		
		while ( !defined $sheetname || !defined $_[0]->worksheet($sheetname) ) {
			print "Choose one sheet from the above existed sheet: ";
			$sheetname = <STDIN>;
			chomp($sheetname);	
		}
		$worksheet = $_[0]->worksheet( $sheetname );
	}	

	print "The present reference sheet is \"", $worksheet->get_name(),"\"\n";
	our ( $row_min, $row_max ) = $worksheet->row_range();
	our ( $col_min, $col_max ) = $worksheet->col_range();
	
	return 1;
}

sub setDefaultTitle {
	my @title;
	if ( $_[0] =~ /^PCIe$/i ) {
		@title = ("Lane","XNet Names","Card Probe Location","Single Ended Vlow", 
			"Single Ended Vhigh","Digital VEYE","Digital HEYE","Scope VEYE",
			"Scope HEYE");
	} elsif( $_[0] =~ /^SAS$/i ) {
		@title = ("Lane","XNet Names","Card Probe Location","Single Ended Vlow", 
			"Single Ended Vhigh","Digital VEYE","Digital HEYE","Scope VEYE",
			"Scope HEYE");
	} elsif( $_[0] =~ /^FSI$/i ) {
		@title = ("FSI Bus","Xnet Names","Frequency","Single Ended Vlow",
			"Single Ended Vhigh","Rise Time","Rise Time","Setup Time","Hold Time",
			"Period Jitter","Cycle to Cycle Jitter","TIE Jitter");

	} elsif( $_[0] =~ /^PSI$/i ) {
		@title = ("PSI Bus","Xnet Names","Frequency","Single Ended Vlow","Single Ended Vhigh",
			"Differential Vlow","Differential Vhigh","Rise Time","Setup Time","Hold Time",
			"Period Jitter","Cycle to Cycle Jitter","TIE Jitter");
	} elsif( $_[0] =~ /^I2C$/i ) {	
		@title = ("I2C Bus","Net Names","Bit Rate","Vlow","Vhigh","Predicted Rise Time (ns)",
			"Rise Time\n(30%-70% ns)","Monotonic Rising Edge","Fall Time\n(30%-70% ns)",
			"Monotonic Falling Edge","Setup Time","Hold Time");
	} elsif( $_[0] =~ /^Clocks$/i ) {	
		@title = ("Clock Description","XNet Name","Probe Location","Single Ended Vlow",
			"Single Ended Vhigh","Frequency	Period","Rise Time","Fall Time","Duty Cycle",
			"Differential Vlow","Differential Vhigh","Period Jitter","Cycle to Cycle Jitter","TIE Jitter");
	} elsif( $_[0] =~ /^DMI$/i ) { 
		@title = ("DMI","Xnet Names","Frequency","Single Ended Vlow","Single Ended Vhigh","Card Measured VEYE",
			"Card Measured HEYE","Rise Time","Rise Time","Period Jitter","Cycle to Cycle Jitter","TIE Jitter");
	} elsif( $_[0] =~ /^FSP Ethernet$/i ) {
		@title = ("MDIO Bus","Xnet Names","Frequency","Single Ended Vlow",
			"Single Ended Vhigh","Rise Time","Rise Time","Setup Time",
			"Hold Time","Period Jitter","Cycle to Cycle Jitter","TIE Jitter");
	} else { @title = (); }
	return \@title;
}

sub setMatchPatt {
	my $input = <STDIN>;
	my @p = split(/\s+/i, $input) ;
	return \@p;
}

sub findMatch {
	my $rowcount = 0;
	my @matchArr;
	my %matchHash;
	my $titleRow;
	my $netNameCol; 
	my @titleColArr;
	my @matchElementArr;
	my $pattern;
	
	for my $row ( $row_min..$row_max ) {
		for my $col ( $col_min..$col_max ) {
			my $cell = $worksheet->get_cell( $row, $col );
			next unless $cell;
			if ( $cell->value() =~ /net name/i ) {
				$titleRow = $row; 
				$netNameCol = $col;

			} 
			last if ( defined($titleRow) );
		}
		last if ( defined($titleRow) );
	}
	
	my $colbuf = $netNameCol+1;
	my $cell = $worksheet->get_cell($titleRow, $colbuf);
	while ($cell) {
		print $colbuf, " ", $cell->value(),"\n";
		$cell = $worksheet->get_cell($titleRow, ++$colbuf);
	}
	print "Select mutiple columns to export: ";
	chomp ( my $exportCol = <STDIN> );
	@titleColArr = split(/\s+/i, $exportCol) ;
	
	my @exportTitleItem;
	for(@titleColArr) {
		$cell = $worksheet->get_cell($titleRow, $_);
		push @exportTitleItem, $cell->value();
	}

	my $pattArrRef = $_[0];
	if ( $#{$pattArrRef} == 0 ) { $pattern = ${$pattArrRef}[0]; 
	} else {
		my $pattArr = join "){1}.+(", @{$pattArrRef};
		$pattern = "[_]?($pattArr){1}" ;
	}

	my $i = 0;	
	for my $row ( $titleRow+1..$row_max ) { 
		my $cell = $worksheet->get_cell( $row, $netNameCol );
		next unless $cell;	
		if ( $cell->value() =~ /$pattern/i) { 
			${$matchArr[$i]}->[0] = $cell->value();
			for (@titleColArr) { 
				my $cell = $worksheet->get_cell( $row, $_);
				push @{$matchElementArr[$i]}, $cell->value();
			}
			${$matchArr[$i]}->[1] = \@{$matchElementArr[$i]};
			$i++;
		} 
	}
	
	@matchArr = sort { ${$a}->[0] cmp ${$b}->[0] } @matchArr;
	
	my $count = 1;
	for (@matchArr) {
		print $count, ": ", ${$_}->[0], "\n";
		$count++;
	}

	return (\@matchArr, \@exportTitleItem);
}


sub getTitleType() {

	BOOL_TITLE:
	my $title_variable;
	print "Apply a default template? (y/n) ";
	chomp( my $title_bool = <STDIN> );

	if ( $title_bool eq "y" || $title_bool eq "Y" ) { 
	 	
		DEFAULT_TITLE:
		print "\t1)PCIe template\n"; 
		print "\t2)SAS template\n";
		print "\t3)FSI template\n";
		print "\t4)PSI template\n";
		print "\t5)I2C template\n";
		print "\t6)Clocks template\n";
		print "\t7)DMI template\n";
		print "\t8)FSP Ethernet template\n";

		chomp( my $title_chs =<STDIN> );
		
		if ($title_chs == 1 ) { $title_variable = "PCIe";
		} elsif ( $title_chs == 2 ) { $title_variable = "SAS";
		} elsif ( $title_chs == 3 ) { $title_variable = "FSI";
		} elsif ( $title_chs == 4 ) { $title_variable = "PSI";
		} elsif ( $title_chs == 5 ) { $title_variable = "I2C";
		} elsif ( $title_chs == 6 ) { $title_variable = "Clocks";
		} elsif ( $title_chs == 7 ) { $title_variable = "DMI";
		} elsif ( $title_chs == 8 ) { $title_variable = "FSP Ethernet";
		} else { goto DEFAULT_TITLE; }
			
	} elsif ( $title_bool eq "n" || $title_bool eq "N" ) {
		$title_variable = "N";		
	} else { goto BOOL_TITLE; }

	return $title_variable;
}

sub exportXls {
	my $writeSheetName;

	my $file = $_[0];
	my $arrRef = $_[1];
	my $title_variable = $_[2];
	my $colOffset;
	my $titleItemRef = $_[3];
	
	SHEET_NAME:
	print "Name of the exported sheet: ";
	chomp( $writeSheetName = <STDIN> );
	
	if ( $title_variable =~ /^PCIe$/ || $title_variable =~ /^SAS$/ ) {
		$colOffset = 10;
	} elsif ( $title_variable =~ /^FSP Ethernet$/ || $title_variable =~ /^FSI$/ || 
		$title_variable =~ /^I2C$/ || $title_variable =~ /^DMI$/ ) { 
			$colOffset = 13 ;
	} elsif ( $title_variable =~ /^PSI$/ ) {
		$colOffset = 14;
	} elsif ( $title_variable =~ /^Clocks$/ ) {
		$colOffset = 15;
	} else { $colOffset = 1;} 

	unless ( -e $_[0]) {
		my $writebook = Spreadsheet::WriteExcel->new($_[0]); 
		$writebook->close();
		print "build!\n";
	}
	
	if ( -e $_[0] && isDefinedWriteSheet($_[0], $writeSheetName) ) {
		
		SHEET_EDIT:
		print "Sheet name \"", $writeSheetName, "\" existed.\nAppend the previous sheet? (y/n) ";
		my $switch = <STDIN>;
		chomp ($switch);
		
		my $row=0;
		if ( $switch eq "y" || $switch eq "Y" ) { 
			writeToXls($file, $writeSheetName, $arrRef, $switch, $titleItemRef, $title_variable, $colOffset);

		} elsif ( $switch eq "n" || $switch eq "N" ) {
			goto SHEET_NAME;
			
		} else { goto SHEET_EDIT; }
	
	} else {
		my $titlerow;
		my $exportitemcol;
		my $contentrow;
		my $contentcol;
		my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
		my $workbook = $parser->Parse($_[0]);
		if ( !defined $workbook ) { die $parser->error(), ".\n"; }

		my $worksheet = $workbook->AddWorksheet($writeSheetName);
	
		$titlerow=0; 
	
		if ( $title_variable =~ /(pcie)|(sas)|(fsi)|(psi)|(i2c)|(clocks)|(dmi)|(fsp ethernet)/i ) { 
			$contentcol = 1;
			$exportitemcol = $colOffset;
		} else { 
			$contentcol = 0; 
			$exportitemcol = $colOffset;
		}
	
		my $temp = 0;
		foreach( @{&setDefaultTitle($title_variable)} ) {
			$worksheet->AddCell($titlerow, $temp, $_);
			$temp++;
		}

		my $i = 0;
		for ( @{$titleItemRef} ) {
			$worksheet->AddCell($titlerow, $exportitemcol+$i, $_);
			++$i;
		}

		$contentrow = $titlerow + 1;
	
		foreach( @{$arrRef} ) {
			$worksheet->AddCell($contentrow, $contentcol, ${$_}->[0]);

			my $i = 0;
			for ( @{${$_}->[1]} ) { 
				$worksheet->AddCell($contentrow, $exportitemcol+$i , $_);	
				++$i;
			}
			$contentrow++;
		}
		$workbook->SaveAs($_[0]);
	}
}

sub clearSheet {
	my $sheetPtr = &getSheetRange($_[0], $_[1]);	
	my @sheetRange = @{$sheetPtr}; 
	print $_[0];
	my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
	my $template = $parser->Parse($_[0]);
	
	my $tempsheet = $template->worksheet($_[1]);
	
	print join " ",@sheetRange;
	print "\n";
	for my $i ($sheetRange[0]..$sheetRange[1]) {
		for my $j ($sheetRange[2]..$sheetRange[3]) {
			$tempsheet->AddCell( $i, $j, '1');
		}
	}
}

sub writeToXls {
	my $titlerow;
	my $contentrow;
	my $exportitemcol;
	my $contentcol;
	my $resultArrRef = $_[2];
	my $titleItemRef = $_[4];
	my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
	my $workbook = $parser->Parse($_[0]);
	if ( !defined $workbook ) { die $parser->error(), ".\n"; }

	my $worksheet = $workbook->worksheet($_[1]);
	
	if ( $_[3] eq "y" || $_[3] eq "Y" ) {
		my $sheetRange = &getSheetRange($_[0], $_[1]);
		$titlerow = ${$sheetRange}[1]+2;
	} else { $titlerow=0; }
	
	if ( $_[5] =~ /(pcie)|(sas)/i || $_[5] =~ /fsp/i ) { 
		$contentcol = 1;
		$exportitemcol = $_[6];
	} else { 
		$contentcol = 0; 
		$exportitemcol = $_[6];
	}

	my $temp = 0;
	foreach( @{&setDefaultTitle($_[5])} ) {
		$worksheet->AddCell($titlerow, $temp, $_);
		$temp++;
	}
	
	my $i = 0;
	for ( @{$titleItemRef} ) {
		$worksheet->AddCell($titlerow, $exportitemcol+$i, $_);
		++$i;
	}

	$contentrow = $titlerow + 1;

	foreach( @{$resultArrRef} ) {
		$worksheet->AddCell($contentrow, $contentcol, ${$_}->[0]);

		my $i = 0;
		for ( @{${$_}->[1]} ) { 
			$worksheet->AddCell($contentrow, $exportitemcol+$i , $_);	
			++$i;
		}
		$contentrow++;
	}
	$workbook->SaveAs($_[0]);
}

sub getSheetRange {
	my $parser   = Spreadsheet::ParseExcel->new();
	my $workbook = $parser->parse( $_[0] );
	
	if ( !defined $workbook ) { die $parser->error(), ".\n"; }

	my $worksheet = $workbook->worksheet( $_[1] );
	my @sheetRange;
	( $sheetRange[0], $sheetRange[1] ) = $worksheet->row_range();
	( $sheetRange[2], $sheetRange[3] ) = $worksheet->col_range();
	
	return \@sheetRange;
}

sub isDefinedRefSheet {
	return defined $worksheet;
}

sub isDefinedWriteSheet {
	my $tmpparser   = Spreadsheet::ParseExcel->new();
	my $tmpworkbook = $tmpparser->parse( $_[0] );

	if ( !defined $tmpworkbook ) { return 0; }

	my $tmpworksheet = $tmpworkbook->worksheet($_[1]);

	foreach my $sheet ($tmpworkbook->worksheets()) {
		if ( $sheet->get_name() eq $_[1] ) { return 1; }
	}
	return 0;
}

sub offRefFile() {
	undef $workbook;
	undef $worksheet;
	return 1;
}

sub offRefSheet {
	undef $sheetname;
	undef $worksheet;
	return 1;
}

1;
