package MsOffice::Word::Surgeon::Utils;

use Exporter  qw/import/;

our @EXPORT = qw/maybe_preserve_spaces/;

use base 'Exporter';


sub maybe_preserve_spaces {
  my ($txt) = @_;
  return $txt =~ /^\s/ || $txt =~ /\s$/ ? ' xml:space="preserve"' : '';
}

1;



