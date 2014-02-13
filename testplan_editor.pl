#!/usr/bin/perl -w 

use strict;
use warnings;
use loadPI;
use ParseSheet;

use Text::CSV;


while(1) {
	MAIN:
	
	print "\t0) Show all files(.xls/.csv) in present directory\n";
	print "\t1) PI testplan editor: union nets with in a same topology\n";
	print "\t2) SI testplan editor: search matched nets from a referenced sheet\n";
	print "\t3) Off present reference file\n";
	print "\t4) Exit\n";

	chomp(my $chs = <STDIN>);
	
	if ($chs eq "0") { 
		ParseSheet::selectFile();
		print "\n";
	} elsif ($chs eq "1") { 
		my $filesHashPtr= &ParseSheet::selectFile();
		my $chs;
		chomp( $chs = <STDIN>);
		print ${$filesHashPtr}{$chs}, " be seleceted.\n";
		
		my $file = ${$filesHashPtr}{$chs};
		LoadPI::readInputPI($file);
	
		#print "If there is any measured point be named starting from \"L\" character?(y/n) ";
		#my $bool_L = <STDIN>;
		#chomp($bool_L);
		#LoadPI::bool_read_L_exception($bool_L);
		LoadPI::rKeysSet();
		LoadPI::unionIdInit();
		LoadPI::checkConnect();
		LoadPI::extractToExcel($file);

	} elsif ($chs eq "2") {
		my $filename;
		
		if ( !ParseSheet::isDefinedRefSheet() ) {
			SELECTFILE:
			my $filesHashPtr= &ParseSheet::selectFile();
			chomp( my $chs = <STDIN>);			
			if ( $chs > keys(%$filesHashPtr) || $chs <= 0 ) { goto SELECTFILE; }
			else { print "\"${$filesHashPtr}{$chs}\" be seleceted. Loading all sheets name...\n"; }

			$filename = ${$filesHashPtr}{$chs};
			my $workbook = &ParseSheet::openFile($filename);
			ParseSheet::openRefSheet($workbook);	
		} else {
			print "Change reference sheet?(y/n) ";
			my $bool_chgsht = <STDIN>;
			chomp ($bool_chgsht);
			while (1) {
				if ( $bool_chgsht eq "y"|| $bool_chgsht eq "Y" ) { 
					my $filesHashPtr= &ParseSheet::selectFile();
					chomp( my $chs = <STDIN>);
					$filename = ${$filesHashPtr}{$chs};
					ParseSheet::offRefSheet();
					ParseSheet::offRefFile();
					my $workbook = &ParseSheet::openFile($filename);
					ParseSheet::openRefSheet($workbook);
					last;
				} elsif ( $bool_chgsht eq "n" || $bool_chgsht eq "N") { last;}
				else {}
			}
		}

		while (1) {
			print "Set match pattern for search: ";
			my $pattArrRef = &ParseSheet::setMatchPatt();
			#&ParseSheet::findMatch($pattArrRef);
			my ($resultArrRef, $exportTitleRef) = &ParseSheet::findMatch($pattArrRef);
			my $resultNo = scalar @{$resultArrRef};
			print $resultNo, " results founded!\n";
			my $searchPatt = join "( )",@$pattArrRef;
			print "Search pattern is \"", $searchPatt,"\"\n";
			
			OPTION:
			print "Export the search result?(y/n) ";
			my $bool_export;
			chomp( $bool_export = <STDIN> );

			if ( $bool_export eq "y" || $bool_export eq "Y" ) { 
				my $titleType = &ParseSheet::getTitleType();
				ParseSheet::exportXls("output.xls", $resultArrRef, $titleType, $exportTitleRef);
				last;
			} elsif ( $bool_export eq "n" ||$bool_export eq "N" ) { 
		       		print "Reset another search pattern?(y/n) ";
				chomp( my $reset = <STDIN> );
				if ( $reset eq "y" || $reset eq "Y" ) { next; }
				else { last; }
			}
			else { undef $bool_export; goto OPTION; }	
			undef $bool_export;
		}	
		
	} elsif ($chs eq "3") {
		ParseSheet::offRefFile();
	} elsif ($chs eq "4") {
		print "Exit...\n";
		last;
	} else { goto MAIN; }
}
