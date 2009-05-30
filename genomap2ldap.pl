#!/usr/bin/perl

use strict;
use utf8;

use Encode;

use Carp;
use Win32::OLE;
use XML::Twig;
use Net::LDAP;
use Net::LDAP::LDIF;
use Net::LDAP::Entry;

use Getopt::Long qw(:config bundling);
use Pod::Usage;

# Used for debugging:
#use Data::Dumper;

binmode(STDOUT, ':utf8');
binmode(STDERR, ':utf8');

# Command-line arguments:
my ($genomap_file, $ldap_host, $ldap_search_dn, $ldap_search_group_dn, $ldap_bind_dn, $ldap_bind_pass);

pod2usage(-verbose=>0, -exitval=>1) unless GetOptions(
	'help'					=> sub { pod2usage(-verbose=>2, -noperldoc=>1, -exitval=>1); },
	'file|f=s'				=> \$genomap_file,
	'host|h=s'				=> \$ldap_host,
	'search-dn|S=s'			=> \$ldap_search_dn,
	'search-group-dn|G=s'	=> \$ldap_search_group_dn,
	'bind-dn|D=s'			=> \$ldap_bind_dn,
	'bind-pass|w=s'			=> \$ldap_bind_pass,
) && defined $genomap_file;

# Tested on WindowsXP, GenoPro v2.0.1.6, ActivePerl v5.10.0 (Win32::OLE v0.1709, XML::Twig v3.32, Net::LDAP v0.39)

# The list of all individuals (key = individualID)
my %individuals = ();

# The list of all contact informations (key = contactID)
my %contacts = ();

# The list of all groups (with group name as a key)
my %groups = ();

# The list of all DNs in this session
my %dn = ();

######################################
# Helper functions
######################################

sub check_dn($)
{
	local $_ = shift;
	
	if (exists $dn{$_})
	{
		croak('DN ' . $_ . 'already defined in this session. Program cannot process the data reliably, as DN definition mechanizm should be changed.');
	}

	$dn{$_} = 1;
}

sub fix_encoding($)
{
	local $_ = shift;
	s/<\?xml version='1\.0' encoding='unicode'\?>/<?xml version='1.0' encoding='windows-1251'?>/;	
	return $_;
}

sub get_contact_type_weight($)
{
	local $_ = shift;
	
	return 1 unless defined and length > 0;
	return 1 if $_ eq 'Other';
	return 3 if $_ eq 'PrimaryResidence';
	return 2 if $_ eq 'TemporaryResidence';
	return 0 if $_ eq 'WorkPlace';

	croak "The program should never reach here for value '$_'";
}

sub trim($)
{
	local $_ = shift;
	s/^\s*//;
	s/\s*$//;
	return $_; 
}

######################################
# Document loading. XML Parsing
######################################

#my $genoProApp = Win32::OLE->GetActiveObject('GenoPro.Application') or croak "Unable to connect to GenoPro instance. Make sure GenoPro is running.";
#my $genoProApp = Win32::OLE->GetObject(encode("cp1251", '< D:/Documents/документы/контакты/контакты.gno')) or croak "Unable to create GenoPro instance. Make sure GenoPro is installed.";
my $genoProApp = Win32::OLE->GetObject($genomap_file) or croak "Unable to create GenoPro instance. Make sure GenoPro is installed.";

XML::Twig->new(
	twig_handlers => {
		'Individuals/Individual'	=> \&individual,
		'Contacts/Contact'			=> \&contact
	}
)->parse(fix_encoding($genoProApp->GetTextXML()))->purge();

# Sort all contact information by relevance:
while (local (undef, $_) = each %individuals)
{
	my @contacts = sort {
		get_contact_type_weight($b->{'type'}) <=> get_contact_type_weight($a->{'type'})
	} map {
		$contacts{$_} or croak "No contacts found for ID $_";
	} @{$_->{'contacts'}};
	
	$_->{'contacts'} = \@contacts;
	
	foreach my $group (@{$_->{'groups'}})
	{
		push @{$groups{$group}}, $_;
	}
}

######################################
# Loading to LDAP
######################################

my $ldap = undef;
my $ldif = Net::LDAP::LDIF->new(\*STDOUT, 'w', 'onerror' => 'die', 'encode' => 'base64'); # 'change' => 1

if (defined $ldap_host)
{
	$ldap = Net::LDAP->new($ldap_host, 'onerror' => 'die') or croak "$@";
	#$ldap->debug(15);

	print "Connected to host $ldap_host, LDAP v" . $ldap->version() . "\n";

	# Anonymous bind?
	if (defined $ldap_bind_dn && defined $ldap_bind_pass)
	{
		$ldap->bind($ldap_bind_dn, 'password' => $ldap_bind_pass);
	}
	else
	{
		$ldap->bind();
	}
}

while (my ($individual_id, $individual) = each %individuals)
{
	my $entry;

	# Contact information is already sorted by relevance:	
	foreach my $contact (@{$individual->{'contacts'}})
	{
		if (defined $contact->{email})
		{
			if ($contact->{email} !~ /^([-\w]+)\s+([-\w. ]+)\s+<(.*)>$/)
			{
				carp "Unable to parse email $contact->{email} for individual $individual_id";
				next;
			}
			
			unless (defined $entry)
			{
				my $dn = "cn=$1 $2" . (defined $ldap_search_dn ? ',' . $ldap_search_dn : '');
				
				eval {
					$entry = $ldap->search('base' => $dn, 'filter' => '(objectClass=person)', 'scope' => 'base', 'sizelimit' => 1)->pop_entry()
				} if $ldap;

				# DN was not found or we have no LDAP connection:
				unless (defined $entry)
				{
					$entry = Net::LDAP::Entry->new($dn);

					# These attributes only for adding a new entry:
					$entry->add(
						'givenName' => $1,
						'sn' => $2,
						'cn' => "$1 $2",
						'objectClass' => [ qw(inetOrgPerson mozillaAbPersonAlpha) ]
					);
				}
				
				$individual->{'dn'} = $dn;
				
				check_dn($dn);

				# Other attributes for adding or updating the entry:
				$entry->replace('mail' => encode('MIME-Q', $3));
			}
			else
			{
				$entry->replace('mozillaSecondEmail' => encode('MIME-Q', $3)) if !$entry->exists('mozillaSecondEmail');
			}
		}
	}

	unless (defined $entry)
	{
		#carp "Individual $individual_id has no email";
		next;
	}

	# Processing other attributes after the entry has been created:
	
	foreach my $contact (@{$individual->{'contacts'}})
	{
		$entry->replace('telephoneNumber'	=> $contact->{'telephone'})	if defined $contact->{'telephone'}	&& !$entry->exists('telephoneNumber');
		$entry->replace('mobile'			=> $contact->{'mobile'})	if defined $contact->{'mobile'}		&& !$entry->exists('mobile');
		$entry->replace('mozillaHomeUrl'	=> $contact->{'homepage'})	if defined $contact->{'homepage'}	&& !$entry->exists('mozillaHomeUrl');
	}

	$entry->replace('pager'					=> $individual->{'icq'})	if defined $individual->{'icq'};
	
	if ($ldap)
	{
		print(($entry->changetype() eq 'add' ? "Adding " : "Updating ") . $entry->dn() . "\n");
		$entry->update($ldap);
	}
	else
	{
		$ldif->write_entry($entry);
	}
}

while (my ($cn, $group) = each %groups)
{
	my $entry;
	my $dn = 'cn=' . $cn . (defined $ldap_search_group_dn ? ',' . $ldap_search_group_dn : '');

	eval {
		$entry = $ldap->search('base' => $dn, 'filter' => '(objectClass=groupOfNames)', 'scope' => 'base', 'sizelimit' => 1)->pop_entry()
	} if $ldap;

	# DN was not found or we have no LDAP connection:
	unless (defined $entry)
	{ 
		$entry = Net::LDAP::Entry->new($dn);
	
		$entry->add(
			'objectclass' => 'groupOfNames',
			'cn' => $cn
		);
	}

	$entry->replace('member' => [ map { $_->{'dn'} } grep { defined $_->{'dn'} } @{$group} ]) ;
	
	# No members for this group have been added to LDAP:
	next unless scalar($entry->get('member'));
	
	check_dn($dn);
	
	if ($ldap)
	{
		print(($entry->changetype() eq 'add' ? "Adding " : "Updating ") . " group: $dn\n");
	
		$entry->update($ldap);
	}
	else
	{
		$ldif->write_entry($entry);
	}
}

$ldap->unbind() if $ldap;

######################################
# XML callback functions
######################################

sub individual()
{
	my ($twig, $individual_node) = @_;
	
	my $map_name = $individual_node->first_child('Position')->att('GenoMap');
	my $individual_link_id = $individual_node->att('IndividualInternalHyperlink');
	
	# If this individual is actually a link, we add a map name to the list of groups.
	# It is assumed, that linked individuals go in XML later than individuals they refer to
	# (e.g. $individual_link_id > $individual_id).
	if (defined $individual_link_id)
	{
		push @{$individuals{$individual_link_id}->{'groups'}}, $map_name if defined $map_name;
		return;	
	}
	
	my $individual_id = $individual_node->att('ID');
	my $contacts_node = $individual_node->first_child('Contacts');

	if (!defined $contacts_node)
	{
		#carp('No contact information available for individual ' . $individual_id); 
		return;
	}
	
	my @contacts = split /\s*,\s*/, $contacts_node->text();

	if (scalar(@contacts) == 0)
	{
		carp('Contact information list is empty for individual ' . $individual_id);
		return;
	}
	
	$individuals{$individual_id}->{'contacts'} = \@contacts;
	
	if (defined $map_name)
	{
		$individuals{$individual_id}->{'groups'} = [ $map_name ]; 
	}
	
	# Custom tags:
	foreach (qw(Name ICQ Skype JabberId AIM))
	{
		my $property_node = $individual_node->first_child($_);
		$individuals{$individual_id}->{lc $_} = trim($property_node->first_child_text()) if defined $property_node;
	}
}

sub contact()
{
	my ($twig, $contact_node) = @_;

	my %properties = ();

	$contacts{$contact_node->att('ID')} = \%properties;
		
	foreach (qw(Type Email Telephone Mobile Homepage))
	{
		my $property_node = $contact_node->first_child($_);
		$properties{lc $_} = trim($property_node->first_child_text()) if defined $property_node;
	}
}

__END__
=pod

=head1 NAME

genomap2ldap.pl - loads infromation about GenoMap individuals into LDAP directory

=head1 SYNOPSIS

  copy_merger.pl -f genomap.gno [-h ldap.host.com] [-D bind_dn] [-w bind_password] [-S search_dn]
  copy_merger.pl --help

=head1 OPTIONS

=over

=item B<--help>

Prints this help message.

=item B<--host|-h>

Optionally specify the LDAP host to connect to. If not defined, the ldiff output is produced, which can be imported
into LDAP directory 

=item B<--search-dn|-S>

Optionally specify the search DN, which should be used during the lookups for individuals. If this is not defined,
all lookups are made in root DN ("").

=item B<--search-group-dn|-G>

Optionally specify the search DN, which should be used during the lookups for groups. If this is not defined,
all lookups are made in root DN ("").

=item B<--bind-dn|-D>

Optionally specify the bind DN, which should be used during the authentication phase. Note, that both bind DN and bind password
should be specified, otherwise the anonymous bind will take place.

=item B<--bind-pass|-w>

Optionally specify the bind password, which should be used during the authentication phase.

=back

=head1 DESCRIPTION

The program opens map file using GenoMap COM interface. The XML data is parsed and converted to .ldiff format.
The program prints the resulting .ldiff to the screen, if no LDAP host is provided. Otherwise for each added
individual it is checked, if the entry already exists in LDAP directory. If positive, the entry fields are
updated, otherwise a new entry is added.

=cut
