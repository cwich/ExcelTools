#!/usr/bin/perl
use strict;
use warnings;
use Encode qw(decode encode);
use Data::Dumper;
use Spreadsheet::ParseExcel;

my $base_file = shift;
my $diff_file = shift;
my $DEBUG = 0;

my %field_column_names = (
    'Project'                    => 'Project',
    'YP2-ID'                     => 'Key',
    'Artefact-Name'              => 'Summary',
    'BUILD-Responsible'          => 'Build Responsible',
    'CD Exclusion Criterion'     => 'Criterion for exclusion Continuous Delivery (CD)',
    'DEV-Stage Automation Type'  => '[CD-state] development',
    'AC1-Stage Automation Type'  => '[CD-state] integration',
    'PROD-Stage Automation Type' => '[CD-state] production',
    'Cloud Maturity Grade'       => 'Cloud Maturity Grade'
);

my @exclusion_types = (
    'None',
    'End of life',
    'Non-linux middleware',
    'Non-standard system',
    'Currently out of scope'
);
my %exclusion_types = map { $_ => 0 } @exclusion_types;

my @stages = (
    'DEV-Stage Automation Type',
    'AC1-Stage Automation Type',
    'PROD-Stage Automation Type'
);
my %stages = map { $_ => 0 } @stages;

my @automation_types = (
    'None',
    'No automation',
    'Legacy automation',
    'Target automation (ansible)'
);
my %automation_types = map { $_ => 0 } @automation_types;

my $base_file_hashref = parse_excelfile($base_file);
my $diff_file_hashref = parse_excelfile($diff_file);

print "Diffs:\n";

my $remove_count = 0;
foreach my $key (sort keys %{$base_file_hashref}) {
    if (!exists $diff_file_hashref->{$key}) {
        print "  - " . Dumper($base_file_hashref->{$key}) . "\n" if $DEBUG;
        print "  - " . sprint_asset($base_file_hashref->{$key}) . "\n";
        $remove_count++;
    }
}

my $add_count = 0;
foreach my $key (sort keys %{$diff_file_hashref}) {
    if (!exists $base_file_hashref->{$key}) {
        print "  + " . Dumper($diff_file_hashref->{$key}) . "\n" if $DEBUG;
        print "  + " . sprint_asset($diff_file_hashref->{$key}) . "\n";
        $add_count++;
    }
}

my $diff_count = 0;
foreach my $key (sort keys %{$base_file_hashref}) {
    if (exists $diff_file_hashref->{$key}) {
        my $changes =
            asset_diff($base_file_hashref->{$key}, $diff_file_hashref->{$key});
        if (defined $changes && (scalar @$changes) > 0) {
            print "  ~ " . sprint_asset($base_file_hashref->{$key}) . "\n";
            print_diffs($changes);
            $diff_count++;
        }
    }
}
print "\n";

print "Summary:\n";
printf("  removed assets : %3i\n", $remove_count);
printf("  added assets   : %3i\n", $add_count);
printf("  changed assets : %3i\n", $diff_count);

sub sprint_asset {
    my $asset_hashref = shift;
    return sprintf(
        "%-8s '%s' => (%s)",
        $asset_hashref->{'YP2-ID'},
        encode("UTF-8", $asset_hashref->{'Artefact-Name'}),
        encode("UTF-8", $asset_hashref->{'BUILD-Responsible'})
    );
}

sub print_diffs {
    my $diff_arrayref = shift;
    foreach (@$diff_arrayref) {
        print "      " . $_ . "\n";
    }
}

sub asset_diff {
    my $asset1_ref = shift;
    my $asset2_ref = shift;
    my @result;
    my @fields = ();
    push @fields, @stages;
    push @fields, 'CD Exclusion Criterion';

    foreach (@fields) {
        if (exists $asset1_ref->{$_}) {
            if (exists $asset2_ref->{$_}) {
                my $value1 = $asset1_ref->{$_};
                my $value2 = $asset2_ref->{$_};
                if ($value1 ne $value2) {
                    push @result, "$_: '$value1' vs. '$value2'";
                }
            }
            else {
                push @result, "$_: '$asset1_ref->{$_}' vs. missing in file 2";
            }
        }
        else {
            if (exists $asset2_ref->{$_}) {
                push @result, "$_: missing in file 1 vs '$asset2_ref->{$_}'";
            }
            else {
                push @result, "$_: missing in file 1 & 2'";
            }
        }
    }

    return \@result;
}

sub parse_excelfile {

    my $excel_file = shift;

    printf("processing File '%s'\n\n", $excel_file);

    my $parser = Spreadsheet::ParseExcel->new();
    my $workbook = $parser->parse($excel_file);

    if (!defined $workbook) {
        die $parser->error(), ".\n";
    }

    my %file_hash = ();

    for my $worksheet ($workbook->worksheets()) {

        my ($row_min, $row_max) = $worksheet->row_range();
        my ($col_min, $col_max) = $worksheet->col_range();
        print "DEBUG: rows: $row_min - $row_max\n" if $DEBUG;
        print "DEBUG: cols: $col_min - $col_max\n" if $DEBUG;

        my %field_column_position = ();

        for my $col ($col_min .. $col_max) {
            my $cell = $worksheet->get_cell(3, $col);
            next unless $cell;
            my $cell_value = $cell->value();

            foreach my $field (keys %field_column_names) {
                if ($cell_value eq $field_column_names{$field}) {
                    $field_column_position{$field} = $col;
                    printf("DEBUG: found field '%30s' at column %3i\n",
                        $field, $col)
                        if $DEBUG;
                }
            }
        }

        my $artefact_count = 0;
        my %exclusion_count = ();
        foreach (@exclusion_types) {
            $exclusion_count{$_} = 0;
        }
        my %automation_count = ();
        foreach my $stage (@stages) {
            foreach (@automation_types) {
                $automation_count{$stage}{$_} = 0;
            }
        }

        for my $row ($row_min .. $row_max) {
            my $cell = $worksheet->get_cell($row, $field_column_position{'Project'});
            next unless $cell;
            if ($cell->value() eq "Yellow Pages Assets") {
                $artefact_count++;
                my %asset = ();

                foreach my $field (keys %field_column_position) {
                    next if $field eq "Project";

                    my $cell = $worksheet->get_cell($row, $field_column_position{$field});
                    my $value = "";
                    if (defined($cell)) {
                        $value = $cell->value();
                    }
                    if (exists $stages{$field} || $field eq 'CD Exclusion Criterion') {
                        if ($value eq "") {
                            $value = "None";
                        }
                    }
                    $asset{$field} = $value;
                    printf("DEBUG: field '%30s' value '%s'\n", $field, $value) if $DEBUG;
                }

                if (!exists $exclusion_types{$asset{'CD Exclusion Criterion'}}) {
                    printf("INCONSISTENCY: unknown CD exclusion criterion (%s) for asset %-6s\n",
                        $asset{'CD Exclusion Criterion'}, $asset{'YP2-ID'})
                }
                $exclusion_count{$asset{'CD Exclusion Criterion'}}++;
                if ($asset{'CD Exclusion Criterion'} eq "None") {
                    foreach my $stage (@stages) {
                        if (exists $asset{$stage}) {
                            my $value = $asset{$stage};
                            $automation_count{$stage}{$value}++;
                            printf("DEBUG: %30s: '%s' ++\n", $stage, $asset{$stage}) if $DEBUG;
                            if (!exists $automation_types{$asset{$stage}}) {
                                 printf("INCONSISTENCY: unknown automation type (%s) for asset %-6s\n",
                                     $asset{$stage}, $asset{'YP2-ID'})
                            }
                        }
                    }
                }
                print "\n" if $DEBUG;

                $file_hash{ $asset{'YP2-ID'} } = \%asset;
            }
        }

        #print Dumper(\%file_hash);

        printf("  %-33s: %3i\n\n", "Artefact count", $artefact_count);

        print "  CD Exclusion Criterions\n";
        my $exclsum = 0;
        foreach my $exclusion_type (@exclusion_types) {
            my $value = $exclusion_count{$exclusion_type};
            $exclsum += $value;
            printf("    %-30s %3i\n", $exclusion_type, $value);
        }
        if ($exclsum != $artefact_count) {
            printf("INCONSISTENCY: sum of exclusion criterions (%i) does not equal artefact count (%i)",
                $exclsum, $artefact_count);
        }
        print "\n";

        foreach my $stage (@stages) {
            print "  $stage\n";
            my $sum = 0;
            foreach my $autotype (@automation_types) {
                if ($autotype ne "None") {
                    my $value = $automation_count{$stage}{$autotype};
                    $sum += $value;
                    printf("    %-30s %3i\n", $autotype, $value);
                }
            }
            printf("    %30s %3i\n\n", "Sum", $sum);
        }
    }
    return \%file_hash;
}
