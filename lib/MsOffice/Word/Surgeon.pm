package MsOffice::Word::Surgeon;
use 5.10.0;
use Moose;
use MooseX::StrictConstructor;
use Archive::Zip                          qw(AZ_OK);
use Encode                                qw(encode_utf8 decode_utf8);
use Carp                                  qw(croak);
use MsOffice::Word::Surgeon::Run;
use MsOffice::Word::Surgeon::Text;
use MsOffice::Word::Surgeon::Revision;
use MsOffice::Word::Surgeon::PackagePart;

use namespace::clean -except => 'meta';

our $VERSION = '1.08';

has 'docx'          => (is => 'ro', isa => 'Str', required => 1);

has 'zip'           => (is => 'ro', isa => 'Archive::Zip', init_arg => undef,
                        builder => '_zip',   lazy => 1);

has 'parts'         => (is => 'ro', isa => 'HashRef[MsOffice::Word::Surgeon::PackagePart]', init_arg => undef,
                        builder => '_parts', lazy => 1);

has 'document'      => (is => 'ro', isa => 'MsOffice::Word::Surgeon::PackagePart', init_arg => undef,
                        builder => '_document', lazy => 1,
                        handles => [qw/contents original_contents indented_contents plain_text replace/]
                       );



has 'next_rev_id'   => (is => 'bare', isa => 'Num', default => 1, init_arg => undef);
   # used by the PackagePart::revision() method for creating *::Revision objects -- each instance
   # gets a fresh value



#======================================================================
# BUILDING
#======================================================================


# syntactic sugar for supporting ->new($path) instead of ->new(docx => $path)
around BUILDARGS => sub {
  my $orig  = shift;
  my $class = shift;

  if ( @_ == 1 && !ref $_[0] ) {
    return $class->$orig(docx => $_[0]);
  }
  else {
    return $class->$orig(@_);
  }
};



#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _zip {
  my $self = shift;

  my $zip = Archive::Zip->new;
  $zip->read($self->{docx}) == AZ_OK
      or die "cannot unzip $self->{docx}";

  return $zip;
}


sub _parts {
  my $self = shift;

  my $xml     = $self->_content_types;

  my @headers = $xml =~ m[PartName="/word/(header\d+).xml"]g;
  my @footers = $xml =~ m[PartName="/word/(footer\d+).xml"]g;

  my %parts = map {$_ => MsOffice::Word::Surgeon::PackagePart->new(surgeon   => $self,
                                                                   part_name => $_)}
                  ('document', @headers, @footers);

  return \%parts;

  # THINK : headers and footers are also listed in word/_rels/document.xml.rels
  # Should we take this source instead of '[Content_Types].xml' ?
}


sub _document {shift->part('document')}


#======================================================================
# methods
#======================================================================

sub _content_types {
  my ($self, $new_content_types) = @_;
  return $self->xml_member('[Content_Types].xml', $new_content_types);
}



sub part {
  my ($self, $part_name) = @_;
  my $part = $self->parts->{$part_name}
    or die "no such part : $part_name";
  return $part;
}



sub xml_member {
  my ($self, $member_name, $new_content) = @_;

  if (! defined $new_content) {
    my $bytes = $self->zip->contents($member_name)
      or die "no zip member for $member_name";
    return decode_utf8($bytes);
  }
  else {
    my $bytes = encode_utf8($new_content);
    return $self->zip->contents($member_name, $bytes);
  }
}


sub headers {
  my ($self) = @_;
  return sort {substr($a, 6) <=> substr($b, 6)} grep {/^header/} keys $self->parts->%*;
}

sub footers {
  my ($self) = @_;
  return sort {substr($a, 6) <=> substr($b, 6)} grep {/^footer/} keys $self->parts->%*;
}

sub new_rev_id {
  my ($self) = @_;
  return $self->{next_rev_id}++;
}



#======================================================================
# METHODS PROPAGATED TO ALL PARTS
#======================================================================


sub all_parts_do {
  my ($self, $method_name, @args) = @_;

  my $parts = $self->parts;

  my %result;

  $result{$_} = $parts->{$_}->$method_name(@args) for keys %$parts;
  return \%result;
}



#======================================================================
# SAVING THE FILE
#======================================================================


sub overwrite {
  my $self = shift;

  $_->_update_contents_in_zip foreach values $self->parts->%*;
  $self->zip->overwrite;
}

sub save_as {
  my ($self, $docx) = @_;

  $_->_update_contents_in_zip foreach values $self->parts->%*;
  $self->zip->writeToFileNamed($docx) == AZ_OK
    or die "error writing zip archive to $docx";
}


#======================================================================
# DELEGATION TO OTHER CLASSES
#======================================================================

sub revision {
  my $self = shift;

  my $revision = MsOffice::Word::Surgeon::Revision->new(rev_id => $self->surgeon->new_rev_id, @_);
  return $revision->as_xml;
}



1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon - tamper wit the guts of Microsoft docx documents

=head1 SYNOPSIS

  my $surgeon = MsOffice::Word::Surgeon->new(docx => $filename);

  # extract plain text
  my $text = $surgeon->document->plain_text;

  # anonymize
  my %alias = ('Claudio MONTEVERDI' => 'A_____', 'Heinrich SCHÜTZ' => 'B_____');
  my $pattern = join "|", keys %alias;
  my $replacement_callback = sub {
    my %args =  @_;
    my $replacement = $surgeon->revision(to_delete  => $args{matched},
                                       to_insert  => $alias{$args{matched}},
                                       run        => $args{run},
                                       xml_before => $args{xml_before},
                                      );
    return $replacement;
  };
  $surgeon->document->replace(qr[$pattern], $replacement_callback);

  # save the result
  $surgeon->overwrite; # or ->save_as($new_filename);


=head1 DESCRIPTION

=head2 Purpose

This module supports a few operations for modifying or extracting text
from Microsoft Word documents in '.docx' format -- therefore the name
'surgeon'. Since a surgeon does not give life, there is no support for
creating fresh documents; if you have such needs, use one of the other
packages listed in the L<SEE ALSO> section.

Some applications for this module are :

=over

=item *

content extraction in plain text format;

=item *

unlinking fields (equivalent of performing Ctrl-Shift-F9 on the whole document)

=item *

regex replacements within text, for example for :

=over

=item *

anonymization, i.e. replacement of names or adresses by aliases;

=item *

templating, i.e. replacement of special markup by contents coming from a data tree
(see also L<MsOffice::Word::Template>).

=back

=item *

pretty-printing the internal XML structure

=back

=head2 Operating mode

The format of Microsoft C<.docx> documents is described in
L<http://www.ecma-international.org/publications/standards/Ecma-376.htm>
and  L<http://officeopenxml.com/>. An excellent introduction can be
found at L<https://www.toptal.com/xml/an-informal-introduction-to-docx>.
Internally, a document is a zipped
archive, where the member named C<word/document.xml> stores the main
document contents, in XML format.

The present module does not parse all details of the whole XML
structure because it only focuses on I<text> nodes (those that contain
literal text) and I<run> nodes (those that contain text formatting
properties). All remaining XML information, for example for
representing sections, paragraphs, tables, etc., is stored as opaque
XML fragments; these fragments are re-inserted at proper places when
reassembling the whole document after having modified some text nodes.


=head1 METHODS

=head2 Constructor

=head3 new

  my $surgeon = MsOffice::Word::Surgeon->new(docx => $filename);
  # or simply : ->new($filename);

Builds a new surgeon instance, initialized with the contents of the given filename.

=head2 Accessors

=head3 docx

Path to the C<.docx> file

=head3 zip

Instance of L<Archive::Zip> associated with this file

=head3 parts

Hashref to L<MsOffice::Word::Surgeon::PackagePart> objects, keyed by their partname in the ZIP file.

=head3 document

Shortcut to C<< $surgeon->part('document') >> -- the 
L<MsOffice::Word::Surgeon::PackagePart> object corresponding to the main document.
See the C<PackagePart> documentation for operations on part objects.



=head3 headers

  my @header_parts = $surgeon->headers;

Returns the ordered list of names of header members stored in the ZIP file.

=head3 footers

  my @footer_parts = $surgeon->footers;

Returns the ordered list of names of footer members stored in the ZIP file.



=head2 Other methods

=head3 part

  my $part = $surgeon->part($part_name);

Returns the L<MsOffice::Word::Surgeon::PackagePart> object corresponding to the given part name.


=head3 xml_member

  my $xml = $surgeon->xml_member($member_name);
  # or
  $surgeon->xml_member($member_name, $new_xml);

Reads or writes the given member name in the ZIP file, with appropriate utf8 decoding or encoding.


=head3 all_parts_do

  my $result = $surgeon->all_parts_do($method_name => %args);

Calls the given method on all part objects. Results are accumulated
in a hash, with part names as keys to the results.

=head3 save_as

  $surgeon->save_as($docx_file);

Writes the ZIP archive into the given file.


=head3 overwrite

  $surgeon->overwrite;

Writes the updated ZIP archive into the initial file.

=head3 revision

  my $xml = $surgeon->revision(
    to_delete   => $text_to_delete,
    to_insert   => $text_to_insert,
    author      => $author_string,
    date        => $date_string,
    run         => $run_object,
    xml_before  => $xml_string,
  );

This method is syntactic sugar for using the
L<MsOffice::Word::Surgeon::Revision> class.
It generates markup for MsWord revisions (a.k.a. "tracked changes"). Users can
then manually review those revisions within MsWord and accept or reject
them. This is best used in collaboration with the L</replace> method :
the replacement callback can call C<< $self->revision(...) >> to
generate revision marks in the document.

Either C<to_delete> or C<to_insert> (or both) must
be present. Other parameters are optional. The parameters are :

=over

=item to_delete

The string of text to delete (usually this will be the C<matched> argument
passed to the replacement callback).

=item to_insert

The string of new text to insert.

=item author

A short string that will be displayed by MsWord as the "author" of this revision.

=item date

A date (and optional time) in ISO format that will be displayed by
MsWord as the date of this revision. The current date and time
will be used by default.

=item run

A reference to the L<MsOffice::Word::Surgeon::Run> object surrounding
this revision. The formatting properties of that run will be
copied into the C<< <w:r> >> nodes of the deleted and inserted text fragments.


=item xml_before

An optional XML fragment to be inserted before the C<< <w:t> >> node
of the inserted text

=back

This method delegates to the
L<MsOffice::Word::Surgeon::Revision> class for generating the
XML markup.



=head1 SEE ALSO

The L<https://metacpan.org/pod/Document::OOXML> distribution on CPAN
also manipulates C<docx> documents, but with another approach :
internally it uses L<XML::LibXML> and XPath expressions for
manipulating XML nodes. The API has some intersections with the
present module, but there are also some differences : C<Document::OOXML>
has more support for styling, while C<MsOffice::Word::Surgeon>
has more flexible mechanisms for replacing
text fragments.


Other programming languages also have packages for dealing with C<docx> documents; here
are some references :

=over

=item L<https://docs.microsoft.com/en-us/office/open-xml/word-processing>

The C# Open XML SDK from Microsoft

=item L<http://www.ericwhite.com/blog/open-xml-powertools-developer-center/>

Additional functionalities built on top of the XML SDK.

=item L<https://poi.apache.org>

An open source Java library from the Apache foundation.

=item L<https://www.docx4java.org/trac/docx4j>

Another open source Java library, competitor to Apache POI.

=item L<https://phpword.readthedocs.io/en/latest/>

A PHP library dealing not only with Microsoft OOXML documents but also
with OASIS and RTF formats.

=item L<https://pypi.org/project/python-docx/>

A Python library, documented at L<https://python-docx.readthedocs.io/en/latest/>.

=back

As far as I can tell, most of these libraries provide objects and methods that
closely reflect the complete XML structure : for example they have classes for
paragraphes, styles, fonts, inline shapes, etc.

The present module is much simpler but also much more limited : it was optimised
for dealing with the text contents and offers no support for presentation or
paging features. However, it has the rare advantage of providing an API for
regex substitutions within Word documents.

The L<MsOffice::Word::Template> module relies on the present module, together with
the L<Perl Template Toolkit|Template>, to implement a templating system for Word documents.


=head1 AUTHOR

Laurent Dami, E<lt>dami AT cpan DOT org<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2019-2022 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
