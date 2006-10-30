use Test::More;
use FindBin ();

SKIP:{
    eval "use Test::Pod 1.00";
    skip "Test::Pod 1.00 required",1 if $@;
    my $path_to_pods = $FindBin::Bin . '/../blib/lib/Spreadsheet/';
    my @poddirs = ($path_to_pods);
    all_pod_files_ok(all_pod_files(@poddirs));
}

