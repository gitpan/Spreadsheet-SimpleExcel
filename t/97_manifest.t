use Test::More;

SKIP:{
    eval "use Test::CheckManifest 1.0";
    skip 'Test::CheckManifest 1.0 is required',1 if $@;
    ok_manifest();
}