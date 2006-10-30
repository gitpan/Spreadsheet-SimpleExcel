use Test::More tests => 1;

SKIP:{
    eval "use Test::CheckManifest 0.4";
    skip 'Test::CheckManifest 0.4 is required',1 if $@;
    ok_manifest();
}