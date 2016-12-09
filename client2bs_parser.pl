#!/cygdrive/c/Perl64/bin/perl

use strict;
use warnings;

# Variables
my $input_raw;
my $input_wav;
my $output_pcmu;

# Usage
if (not defined ($ARGV[0]) or $ARGV[0] ne 'client2bs.raw') { 
  print("Usage: client2bs_parser.pl client2bs.raw\n\n");
  print("       Generates an output file: client2bs.pcmu\n");
  print("       client2bs.raw is obtained from BS\n");
}
else {
  # Store the RAW file
  open (INPUT_RAW, "<client2bs.raw") || die "DIE: Could not find the input RAW file";
  $input_raw = <INPUT_RAW>;
  close INPUT_RAW;

  # Strip te G.711 headers
  $input_raw =~ s/80  0 .. .. .. .. .. .. .. .. .. ..//g;
  $input_raw =~ s/ //g;

  # Convert the file to ASCII
  $output_pcmu = pack("H*", $input_raw);

  # Write to a file
  open (OUTPUT_PCMU, ">client2bs.pcmu") || die "DIE: Could not open the output PCMU file";
  print OUTPUT_PCMU $output_pcmu; #DEBUG
  close OUTPUT_PCMU;
}

