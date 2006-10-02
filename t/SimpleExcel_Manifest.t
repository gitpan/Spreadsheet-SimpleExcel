use Test::More tests => 1;

SKIP:{
    eval "use Test::CheckManifest";
    skip 'Test::CheckManifest is required',1 if $@;
    ok_manifest();
}