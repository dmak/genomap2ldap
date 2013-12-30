#!/usr/bin/perl
#
# Transforms URLs for individuals into tags.
#
use strict;
use utf8;

use Archive::Zip qw(:ERROR_CODES :CONSTANTS);

use Encode;
use XML::Twig;

use Getopt::Long qw(:config bundling);
use Pod::Usage;

my %url_mappings = (
	qr{http://([-\w]+)\.livejournal\.com}	=> 'LiveJournal',
	qr{http://([-\w]+)\.moikrug\.ru}		=> 'MoiKrug'
);

# Command-line arguments:
my ($genomap_file, $debug);

pod2usage(-verbose=>0, -exitval=>1) unless GetOptions(
	'help'		=> sub { pod2usage(-verbose=>2, -noperldoc=>1, -exitval=>1); },
	'file|f=s'	=> \$genomap_file,
	'verbose'	=> \$debug
) && defined $genomap_file;

my $content_was_modified;

my $genomap_zip = Archive::Zip->new();
die "Unable to read ZIP file $genomap_file" unless $genomap_zip->read($genomap_file) == AZ_OK;
die 'ZIP file is empty' unless $genomap_zip->numberOfMembers();

my $twig = XML::Twig->new(
	twig_handlers => {
		'Contacts/Contact' => \&contact,
	}
);

# All insignificant whitespaces are discarded:
my $genomap_zip_member = ($genomap_zip->members())[0];
$twig->parse($genomap_zip_member->contents());

if ($content_was_modified)
{
	$genomap_zip->contents($genomap_zip_member, encode_utf8($twig->sprint()));
	die "Unable to write ZIP file $genomap_file" unless  $genomap_zip->overwrite() == AZ_OK;
}

$twig->purge();

sub contact()
{
	my ($twig, $contact_node) = @_;

	while (my ($url_regex, $tag_name) = each(%url_mappings))
	{
		if ($contact_node->first_child_text('Homepage') =~ /$url_regex/)
		{
			my $tag_value = $1;
			$contact_node->first_child('Homepage')->delete();

			my $id_re = $contact_node->att('ID') . '(,|$)';
			my @contact_ref_nodes = $twig->find_nodes("//Individuals/Individual/Contacts[string() =~ /$id_re/]");

			die 'There should be exactly one parent individual for contact ' . $contact_node->att('ID') unless scalar(@contact_ref_nodes) == 1;

			my $contact_ref_node = $contact_ref_nodes[0];

			# Insert new tag:
			$contact_ref_node->parent()->insert_new_elt($tag_name, $tag_value);

			$content_was_modified = 1;

			if ($contact_node->children_count() == 0 ||
				($contact_node->children_count() == 1 && ($contact_node->children())[0]->tag() eq 'Type'))
			{
				# The contact information is empty:
				local $_ = $contact_ref_node->text();
				s/$id_re//;

				print 'Found parent ' . $contact_ref_node->parent()->att('ID') . ' to update for contact ' . $contact_node->att('ID') . '; text: "' . $contact_ref_nodes[0]->text() . "\" -> \"$_\"\n" if $debug;

				$contact_ref_node->set_text($_);

				$contact_node->delete();
			}
		}
	}
}

__END__
=pod

=head1 NAME

genomap_fix.pl - fixes the contact information for individuals, converting the URLs to given tags.

=head1 SYNOPSIS

  genomap_fix.pl -f genomap.gno [--verbose]
  genomap_fix.pl --help

=head1 OPTIONS

=over

=item B<--help>

Prints this help message.

=item B<--file|-f>

Specify the genomap file.

=item B<--verbose>

Prints debug information.
