package MsOffice::Word::Surgeon::Replacement;
use feature 'state';
use Moose;
use Moose::Util::TypeConstraints;

use POSIX                          qw(strftime);
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces);
use namespace::clean -except => 'meta';

subtype 'Date_ISO',
  as      'Str',
  where   {/\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2})?Z?/},
  message {"$_ is not a date in ISO format yyyy-mm-ddThh:mm:ss"};


has 'original'    => (is => 'ro', isa => 'Str', required => 1);
has 'replacement' => (is => 'ro', isa => 'Str', required => 1);
has 'author'      => (is => 'ro', isa => 'Str'               );
has 'date'        => (is => 'ro', isa => 'Date_ISO', default =>
                        sub {strftime "%Y-%m-%dT%H:%M:%SZ", localtime});




sub as_xml {
  my ($self) = @_;

  state $rev_id = 0;
  $rev_id++;

  my $date   = $self->date;
  my $old    = $self->original;
  my $new    = $self->replacement;
  my $author = $self->author;

  # special attributes for preserving spaces
  my $space_old = maybe_preserve_spaces($old);
  my $space_new = maybe_preserve_spaces($new);

  my $xml = qq{</w:t></w:r>}
          . qq{<w:del w:id="$rev_id" w:author="$author" w:date="$date">}
          . qq{<w:r><w:delText$space_old>$old</w:delText></w:r>}
          . qq{</w:del>}
          . qq{<w:ins w:id="$rev_id" w:author="$author" w:date="$date">}
          . qq{<w:r><w:t$space_new>$new</w:t></w:r>}
          . qq{</w:ins>}
          . qq{<w:r><w:t xml:space="preserve">};

  # NOTE : the last attribute xml:space="preserve" is not necessarily
  # needed, but we can't know because at this stage there is no
  # information about the content of the next run



  # TODO : the inserted <w:r> should copy the properties of the enclosing
  # run.

  return $xml;
}

1;
