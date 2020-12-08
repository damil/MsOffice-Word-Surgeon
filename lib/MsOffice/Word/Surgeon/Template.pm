package MsOffice::Word::Surgeon::Template;
use feature 'state';
use Moose;
use MooseX::StrictConstructor;
use Carp                           qw(croak);
use Template;

use namespace::clean -except => 'meta';


our $VERSION = '1.04';

has 'surgeon'   => (is => 'ro', isa => 'MsOffice::Word::Surgeon', required => 1);
has 'config'    => (is => 'ro', isa => 'HashRef',                 default => sub { {} });

has 'contents'  => (is => 'rw', isa => 'Str',                     init_arg => undef);




sub BUILD {
  my ($self) = @_;

  $self->contents($self->compile_template);
}




sub compile_template {
  my ($self, %options) = @_;

  my $contents = join "", map {$self->template_fragment($_, %options)}  @{$self->surgeon->runs};

  my $rx_para = qr{
    <w:p            [^>]*>
      <w:r          [^>]*>
        <w:t        [^>]*>
          (\[% .*? %\])   (*SKIP)
          <!--TT2green-->
        </w:t>
      </w:r>
      (?: <w:bookmark   [^>]*> )*
    </w:p>
   }sx;

  my $rx_row = qr{
    <w:tr      [^>]*>
      <w:tc    [^>]*>
         (?:<w:tcPr> .*? </w:tcPr> (*SKIP) )?
         $rx_para
      </w:tc>
      (?:<w:tc> .*? </w:tc>   (*SKIP) )*
    </w:tr>
   }sx;


  # paragraphs to be ignored
  $contents =~ s/$rx_row/$1/g;
  $contents =~ s/$rx_para/$1/g;


  return $contents;
}



sub template_fragment {
  my ($self, $run, %options) = @_;

  my $props              = $run->props;
  my $text_color         = $options{text_color}         // "yellow";
  my $instructions_color = $options{instructions_color} // "green";

warn "PROPS: $props\n";


  if ($props =~ s{<w:highlight w:val="($text_color|$instructions_color)"/>}{}) {
    my $col = $1;
    my $xml  = $run->xml_before;
    my $inner_texts = $run->inner_texts;

    if (@$inner_texts) {
      $xml .= "<w:r>";
      $xml .= "<w:rPr>" . $props . "</w:rPr>" if $props;
      $xml .= "<w:t>[% ";
      $xml .= $_->literal_text . "\n" foreach @$inner_texts;
        # NOTE : adding "\n" because end of lines are used by templating modules
      $xml .= " %]<!--TT2$col--></w:t>";
      $xml .= "</w:r>";

      # TODO : factorize code in common  with ->as_xml() method
    }
    warn "XML: $xml\n";

    return $xml;
  }
  else {
    return $run->as_xml;
  }
}





sub process {
  my ($self, %options) = @_;

  my $template = Template->new($self->config)
    or die Template->error(), "\n";

  my $vars = $options{data} // {};
  my $output = "";
  $template->process(\$self->{contents}, $vars, \$output)
    or die $template->error();

  # # remove remaining XML comments
  # $output =~ s[<!--TT2-->][]g;

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

=pod


  # use the contents as a template
  $surgeon->reduce_all_noises;
  $surgeon->merge_runs;
  my $template = $surgeon->compile_template(
        highlights => 'yellow',
        engine     => 'Template', # or Mojo::Template
   );
  my $new_doc = $template->process(data => \%some_data_tree, %other_options);
  $new_doc->save_as($new_doc_filename);

