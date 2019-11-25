#!/usr/bin/perl
use strict;
use warnings;
use Data::Dumper;
use Encode qw(decode encode);
use Spreadsheet::ParseExcel;

my $base_file = shift;
my $diff_file = shift;
my $DEBUG = 0;

# static field name definitions
my $project_name = 'Project';
my $yp2id_name = 'YP2-ID';
my $artefact_name = 'Artefact-Name';
my $buildresp_name = 'BUILD-Responsible';
my $cdexclcrit_name = 'CD Exclusion Criterion';
my $devstage_name = 'DEV-Stage';
my $intstage_name = 'AC1-Stage';
my $prod_stage_name = 'PROD-Stage';
my $cdpipelineauto_name = 'CD Pipeline Automated';
my $cloudmatgrade_name = 'Cloud Maturity Grade';

# field name to excel header name mapping
my %field_column_names = (
    $project_name                => 'Project',
    $yp2id_name                  => 'Key',
    $artefact_name               => 'Summary',
    $buildresp_name              => 'Build Responsible',
    $cdexclcrit_name             => 'Criterion for exclusion Continuous Delivery (CD)',
    $devstage_name               => '[CD-state] development',
    $intstage_name               => '[CD-state] integration',
    $prod_stage_name             => '[CD-state] production',
    $cdpipelineauto_name         => 'Continuous delivery URL',
    $cloudmatgrade_name          => 'Cloud Maturity Grade'
);

# static value name definitions
my $none_name = 'None';
my @exclusion_types = (
    $none_name,
    'End of life',
    'Non-linux middleware',
    'Non-standard system',
    'Currently out of scope'
);
my %exclusion_types = map { $_ => 0 } @exclusion_types;

my @stages = (
    $devstage_name,
    $intstage_name,
    $prod_stage_name
);
my %stages = map { $_ => 0 } @stages;

my @deployment_types = (
    $none_name,
    'No automation',
    'Legacy automation',
    'Target automation (ansible)'
);
my %deployment_types = map { $_ => 0 } @deployment_types;

my @cloud_maturity_grades = (
    $none_name,
    'Level 0 - Virtualized',
    'Level 1 - Loosely Coupled',
    'Level 2 - Abstracted',
    'Level 3 - Adaptive'
);
my %cloud_maturity_grades = map { $_ => 0 } @cloud_maturity_grades;

# parse excel files
my $base_file_hashref = parse_excelfile($base_file);
my $diff_file_hashref = parse_excelfile($diff_file);

# calculate diffs between excel files
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

exit;

## subroutines ##

sub sprint_asset {
    my $asset_hashref = shift;
    return sprintf(
        "%-8s '%s' => (%s)",
        $asset_hashref->{$yp2id_name},
        encode("UTF-8", $asset_hashref->{$artefact_name}),
        encode("UTF-8", $asset_hashref->{$buildresp_name})
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
    push @fields, $cdexclcrit_name;
    push @fields, $cdpipelineauto_name;
    push @fields, $cloudmatgrade_name;

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
        print "\n" if $DEBUG;

        my $artefact_count = 0;
        my %exclusion_count = ();
        foreach (@exclusion_types) {
            $exclusion_count{$_} = 0;
        }
        my %automation_count = ();
        foreach my $stage (@stages) {
            foreach (@deployment_types) {
                $automation_count{$stage}{$_} = 0;
            }
        }
        my $cd_pipeline_count = 0;
        my %cloudmaturity_count = ();
        foreach (@cloud_maturity_grades) {
            $cloudmaturity_count{$_} = 0;
        }

        for my $row ($row_min .. $row_max) {
            my $cell = $worksheet->get_cell($row, $field_column_position{$project_name});
            next unless $cell;
            if ($cell->value() eq "Yellow Pages Assets") {
                $artefact_count++;
                my %asset = ();

                foreach my $field (keys %field_column_position) {
                    next if $field eq $project_name;

                    my $cell = $worksheet->get_cell($row, $field_column_position{$field});
                    my $value = "";
                    if (defined($cell)) {
                        $value = $cell->value();
                    }
                    if (exists $stages{$field} || $field eq $cdexclcrit_name
                        || $field eq $cloudmatgrade_name) {
                        if ($value eq "") {
                            $value = $none_name;
                        }
                    }
                    $asset{$field} = $value;
                    printf("DEBUG: field '%30s' value '%s'\n", $field, $value) if $DEBUG;
                }

                if (exists $exclusion_types{$asset{$cdexclcrit_name}}) {
                    $exclusion_count{$asset{$cdexclcrit_name}}++;
                } else {
                    printf("INCONSISTENCY: unknown CD exclusion criterion (%s) for asset %-6s\n",
                        $asset{$cdexclcrit_name}, $asset{$yp2id_name})
                }
                if ($asset{$cdexclcrit_name} eq $none_name) {
                    foreach my $stage (@stages) {
                        if (exists $asset{$stage}) {
                            my $value = $asset{$stage};
                            $automation_count{$stage}{$value}++;
                            printf("DEBUG: %30s: '%s' ++\n", $stage, $asset{$stage}) if $DEBUG;
                            if (!exists $deployment_types{$asset{$stage}}) {
                                 printf("INCONSISTENCY: unknown automation type (%s) for asset %-6s\n",
                                     $asset{$stage}, $asset{$yp2id_name})
                            }
                        }
                    }
                }
                if (exists $asset{$cdpipelineauto_name} && $asset{$cdpipelineauto_name} ne "") {
                    $cd_pipeline_count++;
                }
                if (exists $cloud_maturity_grades{$asset{$cloudmatgrade_name}}) {
                    $cloudmaturity_count{$asset{$cloudmatgrade_name}}++;
                    if ($asset{$cloudmatgrade_name} eq $none_name) {
                        printf("WARNING: cloud maturity grade not set for asset %-6s\n",
                            $asset{$yp2id_name});
                    }
                } else {
                    printf("INCONSISTENCY: unknown cloud maturity grade (%s) for asset %-6s\n",
                        $asset{$cloudmatgrade_name}, $asset{$yp2id_name})
                }
                print "\n" if $DEBUG;

                $file_hash{ $asset{$yp2id_name} } = \%asset;
            }
        }

        printf("  %-32s %3i\n\n", "Artefact count", $artefact_count);

        print "  ${cdexclcrit_name}s\n";
        my $exclsum = 0;
        foreach (@exclusion_types) {
            my $value = $exclusion_count{$_};
            $exclsum += $value;
            printf("    %-30s %3i\n", $_, $value);
        }
        if ($exclsum != $artefact_count) {
            printf("INCONSISTENCY: sum of exclusion criterions (%i) does not equal artefact count (%i)",
                $exclsum, $artefact_count);
        }
        print "\n";

        foreach my $stage (@stages) {
            print "  $stage Deployment Types\n";
            my $sum = 0;
            foreach (@deployment_types) {
                if ($_ ne $none_name) {
                    my $value = $automation_count{$stage}{$_};
                    $sum += $value;
                    printf("    %-30s %3i\n", $_, $value);
                }
            }
            printf("    %30s %3i\n\n", "Sum", $sum);
        }

        printf("  %-32s %3i\n\n", "CD Pipeline automated", $cd_pipeline_count);

        print "  ${cloudmatgrade_name}s\n";
        my $cmgsum = 0;
        foreach (@cloud_maturity_grades) {
            my $value = $cloudmaturity_count{$_};
            $cmgsum += $value;
            printf("    %-30s %3i\n", $_, $value);
        }
        if ($cmgsum != $artefact_count) {
            printf("INCONSISTENCY: sum of cloud maturity grades (%i) does not equal artefact count (%i)",
                $cmgsum, $artefact_count);
        }
        print "\n";

    }
    return \%file_hash;
}
