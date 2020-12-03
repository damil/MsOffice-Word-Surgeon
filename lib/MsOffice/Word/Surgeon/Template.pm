package MsOffice::Word::Surgeon::Template;
use feature 'state';
use Moose;
use MooseX::StrictConstructor;
use Carp                           qw(croak);
use Template;

use namespace::clean -except => 'meta';


our $VERSION = '1.04';

has 'surgeon'   => (is => 'ro', isa => 'MsOffice::Word::Surgeon', required => 1);

has 'contents'  => (is => 'ro', isa => 'Str',          required => 1);

has 'config'    => (is => 'ro', isa => 'HashRef',      default => sub { {} });


sub process {
  my ($self, %options) = @_;

  my $template = Template->new($self->config)
    or die Template->error(), "\n";

  my $vars = $options{data} // {};
  my $output = "";
  $template->process(\$self->{contents}, $vars, \$output)
    or die $template->error();

  # remove empty paragraphs coming from TT2 instructions like IF, FOREACH
  $output =~ s[<w:p><w:r><w:t><!--TT2--></w:t></w:r></w:p>(?!</w:tc>)][]g;

  # remove remaining XML comments
  $output =~ s[<!--TT2-->][]g;

  my $new_doc = $self->surgeon->meta->clone_object($self->surgeon);
  $new_doc->contents($output);

  return $new_doc;
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon::Template - TODO

=head1 DESCRIPTION


=head1 METHODS

