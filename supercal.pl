#!/usr/bin/perl
use strict;
use warnings;
use Spreadsheet::WriteExcel;

my $i;
my $j;
my $decval_ten;
my $decval_one;
my $total;
my $cool;
my $line;
my $code;
my $out;
my $result;
my $name;
my $unit; # hehe!
my $parse;
my $tag;
my $version = "0.2";
my $time = time;
my $datetime = localtime (time);
my $workbook;
my $worksheet;
my $start_row;
my $start_col;
my $token;
my $format;
my $tool_path = "C:\\CVS\\cel_fi_SW\\cel_fi_tools";
my $xl_limit = "$tool_path\\CalLimits.xls";
my $calexe = "$tool_path\\CalPostProcess.exe";
my $limit = "$tool_path\\CalLimits.txt";
my $wordpad = "C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe";

#x = 0
my @low_cell_rx = qw(12 14 16 18 20);
my @high_cell_rx = qw(23 25 27 29 31);

#x = 1
my @low_cell_tx = qw(35 37 39 41 43);
my @high_cell_tx = qw(45 47 49 51 53);

#x = 2
my @low_unii_rx_port0 = qw(57 59 61 63 65);
my @high_unii_rx_port0 = qw(68 70 72 74 76);

#x = 3
my @low_unii_rx_port1 = qw(79 81 83 85 87);
my @high_unii_rx_port1 = qw(90 92 94 96 98);

#x = 4
my @low_unii_rx_port2 = qw(101 103 105 107 109);
my @high_unii_rx_port2 = qw(112 114 116 118 120);

#x = 5
my @low_unii_tx_port0 = qw(124 126 128 130 132);
my @high_unii_tx_port0 = qw(135 137 139 141 143);

#x = 6
my @low_unii_tx_port1 = qw(146 148 150 152 154);
my @high_unii_tx_port1 = qw(157 159 161 163 165);

#x = 7
my @low_pwr_det = qw(169 171 173 175 177);
my @high_pwr_det = qw(180 182 184 186 188);

my @row;
my @col;
my @field;
my @digits;
my @input_log;
my @input_cal;
my @in_files;
my @output;
my @calerr;
my @sys = qw(Cell_RX_fail Cell_TX_fail UNII_C_RX_fail UNII_D_RX_fail UNII_A_RX_fail UNII_A_TX_fail UNII_B_TX_fail Cell_TX_power_det_fail AFC_fail Quality_metrics_EVM_fail);
my @too = qw(Low High);
my @cell_tx_freq = qw(UARFCN_9622_Cell_TX_meas_fail UARFCN_9686_Cell_TX_meas_fail UARFCN_9750_Cell_TX_meas_fail UARFCN_9814_Cell_TX_meas_fail UARFCN_9878_Cell_TX_meas_fail);
my @cell_rx_freq = qw(UARFCN_10572_Cell_RX_meas_fail UARFCN_10636_Cell_RX_meas_fail UARFCN_10700_Cell_RX_meas_fail UARFCN_10764_Cell_RX_meas_fail UARFCN_10828_Cell_RX_meas_fail);
my @unii_tx_freq = qw(5536MHz_UNII_TX 5568MHz_UNII_TX 5600MHz_UNII_TX 5632MHz_UNII_TX 5664MHz_UNII_TX);
my @unii_rx_freq = qw(5186MHz_UNII_RX 5218MHz_UNII_RX 5250MHz_UNII_RX 5282MHz_UNII_RX 5314MHz_UNII_RX);
my @loopbk = qw(UNII_loopback_C_to_A_fail UNII_loopback_D_to_B_fail Cell_loopback_fail);
my @meas = qw(0-116_for_cell_RX 0-7_cell_pwr_det 0-15_for_UNII_TX_and_RX 0-119_for_cell_TX 0-3_for_AFC);
my @qual = qw(RMS_EVM_% Peak_EVM_% Freq_Err_Hz Origin_Offset_dB Phase_Err_0deg Mag_Err_dB FF-tester_couldnt_perform_the_quality_metrics);

# Parse the CalResults file
open (INPUT_CAL, "type CalCheckRslts.txt |") || die "DIED: INPUT";
@input_cal = <INPUT_CAL>;
close INPUT_CAL;

if ($#input_cal == "-1") {
	print("\n Cannot find CalCheckRslts.txt file - proceeding with file parse\n");
	print("\n Pls select files type: 1-rawclibration or 2-calibrationerrors\n");
	chomp ($parse = <>);

	# Get files
	if ($parse == 1) {
		open (INPUT, "dir *raw* |") || die "DIED: INPUT";
		@input_log = <INPUT>;
		close INPUT;
		$tag = "rawcalibration";
	}
	elsif ($parse == 2) {
		open (INPUT, "dir *errors* |") || die "DIED: INPUT";
		@input_log = <INPUT>;
		close INPUT;
		$tag = "calibrationerrors";
	}
	# Generate cal results
	foreach $line (@input_log) {
		if ($line =~ m/(\d+|\d+_\d+).(raw|calib)/) {
				$name = $1;
				if ($parse == 1) {
					system("$calexe $limit $1.$tag $1.txt 1 > $1.CalCheckRslts ");
					open (IN_FILES, "type $1.CalCheckRslts |") || die "DIED: FILES";
				} 
				elsif ($parse == 2) {
					open (IN_FILES, "type $1.$tag |") || die "DIED: FILES";
				}
				@in_files = <IN_FILES>;
				if ($1 =~ m/^100/) {
					foreach $code (@in_files) {
						if ($code =~ m/(^\d+\D* )/) {
							chomp $code;
							push @output, "\n$code";
							@digits = split(//, $1);
							push @output, "\t\t$sys[$digits[0]]";
							if ($digits[0] == 0) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$cell_rx_freq[$digits[2]]";
								push @output, "\t\t$meas[0]";
							}
							elsif ($digits[0] == 1) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$cell_tx_freq[$digits[2]]";
								push @output, "\t\t$meas[3]";
							}
							elsif (($digits[0] == 2) || ($digits[0] == 3) || ($digits[0] == 4)) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$unii_rx_freq[$digits[2]]";
								push @output, "\t\t$meas[2]";
							}
							elsif (($digits[0] == 5) || ($digits[0] == 6)) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$unii_tx_freq[$digits[2]]";
								push @output, "\t\t$meas[2]";
							}
							elsif ($digits[0] == 7) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$cell_tx_freq[$digits[2]]";
								push @output, "\t\t$meas[1]";
							}
							elsif ($digits[0] == 8) {
								push @output, "\t\t$meas[4]";
							}
							elsif ($digits[0] == 9) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$loopbk[$digits[2]]";
								if($digits[4] eq "F") {
									push @output, "\t$qual[6]";
								}
								else {
									push @output, "\t$qual[$digits[4]]";
								}
							}
						}
					}
					open (RESULTS, ">>$name.CalCheckRslts") || die "DIED: RESULTS";
					foreach $result (@output) {
						print RESULTS $result;
					}
					close RESULTS;
					close IN_FILES;
				}
				elsif ($1 =~ m/^101/) {
					foreach $code (@in_files) {
						if ($code =~ m/(^\d+\D* )/) {
							chomp $code;
							push @output, "\n$code";
							@digits = split(//, $1);
							push @output, "\t\t$sys[$digits[0]]";
							if ($digits[0] == 0) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$cell_rx_freq[$digits[2]]";
								push @output, "\t\t$meas[0]";
							}
							elsif ($digits[0] == 1) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$cell_tx_freq[$digits[2]]";
								push @output, "\t\t$meas[3]";
							}
							elsif (($digits[0] == 2) || ($digits[0] == 3)) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$unii_rx_freq[$digits[2]]";
								push @output, "\t\t$meas[2]";
							}
							elsif (($digits[0] == 5) || ($digits[0] == 6)) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$unii_tx_freq[$digits[2]]";
								push @output, "\t\t$meas[2]";
							}
							elsif ($digits[0] == 8) {
								push @output, "\t\t$meas[4]";
							}
							elsif ($digits[0] == 9) {
								push @output, "\t$too[$digits[1]]";
								push @output, "\t$loopbk[$digits[2]]";
								if($digits[4] eq "F") {
									push @output, "\t$qual[6]";
								}
								else {
									push @output, "\t$qual[$digits[4]]";
								}
							}
						}
					}
					open (RESULTS, ">>$name.CalCheckRslts") || die "DIED: RESULTS";
					foreach $result (@output) {
						print RESULTS $result;
					}
					close RESULTS;
					close IN_FILES;
				}
			}
	}
	if ($parse == 1) {
		system("del CalCheckRslts.txt");
	}
}
else {
	print("Is the ChkCalResults or calibrationerrors file from 1-WU or 2-CU?\n");
	chomp ($unit = <>);
	# Generate cal results
	if ($unit == 1) {
		foreach $code (@input_cal) {
			if ($code =~ m/(^\d+\D* )/) {
				chomp $code;
				push @output, "\n$code";
				@digits = split(//, $1);
				push @output, "\t\t$sys[$digits[0]]";
				if ($digits[0] == 0) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$cell_rx_freq[$digits[2]]";
					push @output, "\t\t$meas[0]";
					if ($digits[1] == 0) {
						push @row, "$low_cell_rx[$digits[2]]";
					}
					elsif ($digits[1] == 1) {
						push @row, "$high_cell_rx[$digits[2]]";
					}
				}
				elsif ($digits[0] == 1) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$cell_tx_freq[$digits[2]]";
					push @output, "\t\t$meas[3]";
					if ($digits[1] == 0) {
						push @row, "$low_cell_tx[$digits[2]]";
					}
					elsif ($digits[1] == 1) {
						push @row, "$high_cell_tx[$digits[2]]";
					}
				}
				elsif (($digits[0] == 2) || ($digits[0] == 3) || ($digits[0] == 4)) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$unii_rx_freq[$digits[2]]";
					push @output, "\t\t$meas[2]";
					if (($digits[1] == 0) && ($digits[0] == 2)) {
						push @row, "$low_unii_rx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 2)) {
						push @row, "$high_unii_rx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 0) && ($digits[0] == 3)) {
						push @row, "$low_unii_rx_port1[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 3)) {
						push @row, "$high_unii_rx_port1[$digits[2]]";
					}
					elsif (($digits[1] == 0) && ($digits[0] == 4)) {
						push @row, "$low_unii_rx_port2[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 4)) {
						push @row, "$high_unii_rx_port2[$digits[2]]";
					}
				}
				elsif (($digits[0] == 5) || ($digits[0] == 6)) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$unii_tx_freq[$digits[2]]";
					push @output, "\t\t$meas[2]";
					if (($digits[1] == 0) && ($digits[0] == 5)) {
						push @row, "$low_unii_tx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 5)) {
						push @row, "$high_unii_tx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 0) && ($digits[0] == 6)) {
						push @row, "$low_unii_tx_port1[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 6)) {
						push @row, "$high_unii_tx_port1[$digits[2]]";
					}
				}
				elsif ($digits[0] == 7) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$cell_tx_freq[$digits[2]]";
					push @output, "\t\t$meas[1]";
					if ($digits[1] == 0) {
						push @row, "$low_pwr_det[$digits[2]]";
					}
					elsif ($digits[1] == 1) {
						push @row, "$high_pwr_det[$digits[2]]";
					}
				}
				elsif ($digits[0] == 8) {
					push @output, "\t\t$meas[4]";
				}
				elsif ($digits[0] == 9) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$loopbk[$digits[2]]";
					if(($digits[3] eq 'F') && ($digits[4] eq 'F')) {
						push @output, "\t$qual[6]";
					}
				}
				if (!(($digits[3] eq 'F') && ($digits[4] eq 'F'))) {
					$decval_ten = hex($digits[3]) * 16;
					$decval_one = hex($digits[4]);
					$total = $decval_ten + $decval_one;
					push @col, "$total";
				}
			}
		}
		open (CALFILE, $limit) or die "CAL_FILE";
		$workbook = Spreadsheet::WriteExcel->new($xl_limit);
		$worksheet = $workbook->add_worksheet();

		# Row and column are zero indexed
		$start_row = 0;

		$format = $workbook->add_format();
		$format->set_bg_color('yellow');

#test		print ("@col\n");
#test		print ("@row\n");

		$i = 0;
		$j = 0;
		while (<CALFILE>) {
			chomp;
			# Split on single tab
			@field = split(',', $_);
			$start_col = 0;
			foreach $token (@field) {
				if ((defined $row[$i]) && (defined $col[$j])) {
#test					print ("$row[$i] and $col[$j]\n");
#test					print ("$start_row and $start_col\n");
					if (($start_row == $row[$i]) && ($start_col == $col[$j])) {
						$worksheet->write($start_row, $start_col, $token, $format);
						$j++;
						$i++;
#test						print("^^^^^MATCH^^^^^^\n");
#test						print("^^^^^MATCH^^^^^^\n");
#test						print("^^^^^MATCH^^^^^^\n");
					}
					else {
						$worksheet->write($start_row, $start_col, $token);
					}
				}
				else {
#test					print ("no match $start_row and $start_col\n");
					$worksheet->write($start_row, $start_col, $token);
				}
				$start_col++;
			}
			$start_row++;
		}
		$workbook->close();
	}
	elsif ($unit == 2) {
		foreach $code (@input_cal) {
			if ($code =~ m/(^\d+\D* )/) {
				chomp $code;
				push @output, "\n$code";
				@digits = split(//, $1);
				push @output, "\t\t$sys[$digits[0]]";
				if ($digits[0] == 0) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$cell_rx_freq[$digits[2]]";
					push @output, "\t\t$meas[0]";
					if ($digits[1] == 0) {
						push @row, "$low_cell_rx[$digits[2]]";
					}
					elsif ($digits[1] == 1) {
						push @row, "$high_cell_rx[$digits[2]]";
					}
				}
				elsif ($digits[0] == 1) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$cell_tx_freq[$digits[2]]";
					push @output, "\t\t$meas[3]";
					if ($digits[1] == 0) {
						push @row, "$low_cell_tx[$digits[2]]";
					}
					elsif ($digits[1] == 1) {
						push @row, "$high_cell_tx[$digits[2]]";
					}
				}
				elsif ($digits[0] == 2 || $digits[0] == 3) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$unii_rx_freq[$digits[2]]";
					push @output, "\t\t$meas[2]";
					if (($digits[1] == 0) && ($digits[0] == 2)) {
						push @row, "$low_unii_rx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 2)) {
						push @row, "$high_unii_rx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 0) && ($digits[0] == 3)) {
						push @row, "$low_unii_rx_port1[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 3)) {
						push @row, "$high_unii_rx_port1[$digits[2]]";
					}
				}
				elsif ($digits[0] == 5 || $digits[0] == 6) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$unii_tx_freq[$digits[2]]";
					push @output, "\t\t$meas[2]";
					if (($digits[1] == 0) && ($digits[0] == 5)) {
						push @row, "$low_unii_tx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 5)) {
						push @row, "$high_unii_tx_port0[$digits[2]]";
					}
					elsif (($digits[1] == 0) && ($digits[0] == 6)) {
						push @row, "$low_unii_tx_port1[$digits[2]]";
					}
					elsif (($digits[1] == 1) && ($digits[0] == 6)) {
						push @row, "$high_unii_tx_port1[$digits[2]]";
					}
				}
				elsif ($digits[0] == 8) {
					push @output, "\t\t$meas[4]";

				}
				elsif ($digits[0] == 9) {
					push @output, "\t$too[$digits[1]]";
					push @output, "\t$loopbk[$digits[2]]";
					if(($digits[3] eq 'F') && ($digits[4] eq 'F')) {
							push @output, "\t$qual[6]";
					}
				}
				if (!(($digits[3] eq 'F') && ($digits[4] eq 'F'))) {
					$decval_ten = hex($digits[3]) * 16;
					$decval_one = hex($digits[4]);
					$total = $decval_ten + $decval_one;
					push @col, "$total";
				}
			}
		}
		open (CALFILE, $limit) or die "CAL_FILE";
		$workbook = Spreadsheet::WriteExcel->new($xl_limit);
		$worksheet = $workbook->add_worksheet();

		# Row and column are zero indexed
		$start_row = 0;

		$format = $workbook->add_format();
		$format->set_bg_color('yellow');

#test		print ("@col\n");
#test		print ("@row\n");

		$i = 0;
		$j = 0;
		while (<CALFILE>) {
			chomp;
			# Split on single tab
			@field = split(',', $_);
			$start_col = 0;
			foreach $token (@field) {
				if ((defined $row[$i]) && (defined $col[$j])) {
#test					print ("$row[$i] and $col[$j]\n");
#test					print ("$start_row and $start_col\n");
					if (($start_row == $row[$i]) && ($start_col == $col[$j])) {
						$worksheet->write($start_row, $start_col, $token, $format);
						$j++;
						$i++;
#test						print("^^^^^MATCH^^^^^^\n");
#test						print("^^^^^MATCH^^^^^^\n");
#test						print("^^^^^MATCH^^^^^^\n");
					}
					else {
						$worksheet->write($start_row, $start_col, $token);
					}
				}
				else {
#test					print ("no match $start_row and $start_col\n");
					$worksheet->write($start_row, $start_col, $token);
				}
				$start_col++;
			}
			$start_row++;
		}
		$workbook->close();
	}
	open (RESULTS, ">>SuperCalCheckRslts_${time}.txt") || die "DIED: RESULTS";
	print RESULTS "Supercal v$version  $datetime\n";
	foreach $result (@output) {
		print RESULTS $result;
	}
	close RESULTS;
	close IN_FILES;
	close CALFILE;
	system "start wordpad SuperCalCheckRslts_${time}.txt";
	system "$xl_limit";

	
}