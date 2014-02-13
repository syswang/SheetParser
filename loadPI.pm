package LoadPI;
use strict;
use warnings;
use Text::CSV;
use Spreadsheet::WriteExcel;
use List::MoreUtils qw/uniq/;

our @nets;
our @exceptionL;
our %ExceptLHash;
our %HashOfNets;
 
sub readInputPI {
	## Open input PI file
	my $file = $_[0];
#	chomp($file);
	
	my $checkedCol;
	my $netNameCol;
	our @titleArray;

	my $csv = Text::CSV->new ( { binary => 1 } )  # should set binary attribute.
	or die "Cannot use CSV: ".Text::CSV->error_diag ();
	open my $io, "<", "$file" or die "$file : $!";

	## Read until the title
	while (my $row = $csv->getline ($io)) {
		push @titleArray, $row;
		for( 0..$#{$row} ){
			$netNameCol = ($_) if( ${${row}}[$_] =~ /name/i );
			$checkedCol = ($_) if( ${${row}}[$_] =~ /measured/i );	
		}
		if ( defined($checkedCol) ) { last;}		
	}

	## Read data to %HashOfNets{$netName}  
	our %HashOfNets;  #hash of the measured data
	my @netArray;  #array address of the data in one net 
	my $i = 0;  #index for the @netArray 
	while (my $row = $csv->getline ($io)) {
		if (${$row}[$netNameCol]){
			$i++;
			$HashOfNets{ ${$row}[$netNameCol] } = \@{$netArray[$i]};
		} 
		unless ( !${$row}[$checkedCol] ) {
			push @{$netArray[$i]}, ${$row}[$checkedCol];
		}	
	}
	our @nets = sort keys(%HashOfNets);		#all netName array
}

sub bool_read_L_exception {
	## Read exception of L starting character to Hash
	while (1) {
		if ( $_[0] eq "y" || $_[0] eq "Y" ) { 
			print "debug\n";
			print "Type in your .txt exception \"L\" character starting measurement points file name: ";
			my $file = <STDIN>;
			chomp($file);
			&readExceptionL($file); 
			last;
		}
		elsif ( $_[0] eq "n" || $_[0] eq "N") { last;}
		else { }
	}	
}



our %rnets;	#hash of reverse Keys from idx to nets
our %id;	#hash of root id for union judge


sub checkConnect {
	## check the interconnect 
	our @nets;
	our %HashOfNets; 

	for ( my $idx = 0; $idx < $#nets; $idx++ ){
		for ( 0..$#{$HashOfNets{$nets[$idx]}} ){
			next unless ( ${$HashOfNets{$nets[$idx]}}[$_] =~ /^[RLFQ]/);
			&checkIntcnt( $idx+1, $#nets, ${$HashOfNets{$nets[$idx]}}[$_] );			
		}
	}
}

sub readExceptionL() {
	my $excptIdx = 0;	
	my @exceptTmpArr;
	our @exceptionL;
	our %ExceptLHash;
	open (FILE, $_[0]) or die "$_[0] : $!";
	while (<FILE>) {
		chomp;
		push @exceptionL, $_;
	}
	close (FILE);

	for my $key (@nets) {
		@{$exceptTmpArr[$excptIdx]} = intersect( @{$HashOfNets{$key}}, @exceptionL );
		if ( @{$exceptTmpArr[$excptIdx]} ) {
			$ExceptLHash{$key} = \@{$exceptTmpArr[$excptIdx]} ;
			$excptIdx++;
		}
	}
}

sub rKeysSet() {
	our %rnets;	#hash of reverse Keys from idx to nets
	for (my $i=0; $i <= $#nets; $i++ ) {
		$rnets{ $i } = $nets[$i];		#initialization of %rnets
	}
}

sub unionIdInit() {
	our %id;	#hash of root id for union judge
	for (my $i=0; $i <= $#nets; $i++ ) {
		$id{ $nets[$i] } = $i;		#initialization of %id
	}
}

#sub debugId() {
#	print $_, " => ", $id{$_},"\n" foreach nets %id;
#	print $_, " => ", $nets[$_],"\n" for 0..$#nets;
#	print $_, " => ", $rnets{$_},"\n" foreach nets %rnets;
#}

sub checkIntcnt() {
	for ( my $idx = $_[0]; $idx <= $_[1]; $idx++ ){
		for ( 0 .. $#{$HashOfNets{$nets[$idx]}} ) {
			next unless ( ${$HashOfNets{$nets[$idx]}}[$_] =~ /^[RLFQ]/);		## Condition for R, L, F, Q Comparision 
			if( $_[2] eq ${$HashOfNets{$nets[$idx]}}[$_] ) {	#repeat measured point
				#print "Repeat! in ", $_[2], " & ", ${$HashOfNets{$nets[$idx]}}[$_], " ";
				#print $nets[$idx], " to ", &getKeyByValue( $_[2], \%HashOfNets ), "\n";
				my $host = &getKeyByValue( $_[2], \%HashOfNets );
				&unionId($host, $nets[$idx]);		#connect two %id of nets  
				last;
			}
		}
	}
	return;
}

sub isConnected {
	return &root($_[0]) == &root($_[1]);
}

sub root() {
	our %id;
	my $tmp = $_[0] ;
	my $i = 0;
	while ( !( $tmp eq $nets[$i] ) ) { $i++; } 
	while ( $i != $id{ $nets[$i] } ) { 
		$i = $id{ $nets[$i] }; 
	}
	return $i;
}

sub unionId() {
	our %id;
	my $p = &root( $_[0] );
	my $q = &root( $_[1] );
	if ( $p == $q ) { return; }
	$id{ $nets[$q] } = $p;
}

sub getKeyByValue() {
	for my $key (@nets){
		for ( 0 .. $#{$_[1]{$key}} ) {
			if ( $_[0] eq ${$_[1]{$key}}[$_] ) { return $key; }
		}
	}
}

sub checkException() {
	our %ExceptLHash;
	if (exists $ExceptLHash{ $_[0] } ) {
		push $_[1], @{$ExceptLHash{ $_[0] }};
		delete $ExceptLHash{ $_[0] };	
	}
}

sub extractToExcel {
	## Export result to .xls file
	my $tmpFile;
	if ( $_[0] =~ /(.+)\.(.+)/) { $tmpFile = $1; }

	my $workbook = Spreadsheet::WriteExcel->new("$tmpFile.xls");
	my $worksheet = $workbook->add_worksheet();
	my $netRow = 0;
	my $measuredRow = 0;
	my @buffer;
	
	our %rnets;
	our @titleArray;

	print "The extraction result named: $tmpFile.xls\n";
	
	$worksheet->write_col( $netRow, 0, \@titleArray);
	$netRow += $#titleArray+1;
	$measuredRow += $#titleArray+1;

	for (my $i = 0; $i<=$#nets; $i++) {
		#write first net
		next unless $rnets{$i};
		$worksheet->write( $netRow, 0, $nets[$i] );
		++$netRow;
		push @buffer, grep( !/^[CDRLFQP]/, @{$HashOfNets{$nets[$i]}} );
		&checkException( $nets[$i], \@buffer);
		#$worksheet->write_col( $measuredRow, 1, \@buffer );
		#$measuredRow += $#buffer+1;
		delete $rnets{$i};

		#scan and write same id net
		for (my $j = $i+1; $j<=$#nets; $j++) {
			next unless $rnets{$j};
			if ( isConnected($nets[$i], $nets[$j]) ) { 
				$worksheet->write( $netRow, 0, $rnets{$j} );
				++$netRow;
				push @buffer, grep( !/^[CDRLFQP]/ ,@{$HashOfNets{$nets[$j]}} );
				&checkException( $nets[$j], \@buffer);
				#$worksheet->write_col( $measuredRow, 1, \@buffer);
				#$measuredRow += $#buffer+1;
				delete $rnets{$j};
			}
		}
		
		my @uniqBuffer = sort { $a cmp $b } uniq @buffer;
		$worksheet->write_col( $measuredRow, 1, \@uniqBuffer);
		$measuredRow += $#uniqBuffer+1;
		
		#the next start Row preparation
		if ( $measuredRow >= $netRow ) { 
			$measuredRow += 1;
			$netRow = $measuredRow;
		}
		else { 
			$netRow += 1; 
			$measuredRow = $netRow;	
		}
		@buffer = ();
	}
}

1;
