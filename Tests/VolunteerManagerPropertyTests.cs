using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for VolunteerManager class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 1.4, 7.1
    /// </summary>
    [TestFixture]
    public class VolunteerManagerPropertyTests
    {
        private VolunteerManager _volunteerManager = null!;

        [SetUp]
        public void Setup()
        {
            _volunteerManager = new VolunteerManager();
        }

        /// <summary>
        /// Custom generator for valid email addresses
        /// </summary>
        private static Gen<string> ValidEmailGen()
        {
            // Generate local part: must start with alphanumeric, can contain ._- in middle
            var localPartGen = from firstChar in Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray())
                              from length in Gen.Choose(0, 19)
                              from middleChars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789._-".ToCharArray()))
                              select firstChar + new string(middleChars);

            // Generate domain part: must start with alphanumeric, can contain - in middle
            var domainGen = from firstChar in Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray())
                           from length in Gen.Choose(1, 14)
                           from middleChars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789-".ToCharArray()))
                           select firstChar + new string(middleChars);

            var tldGen = Gen.Elements("com", "org", "net", "it", "edu");

            return from local in localPartGen
                   from domain in domainGen
                   from tld in tldGen
                   select $"{local}@{domain}.{tld}";
        }

        /// <summary>
        /// Custom generator for valid volunteer surnames (non-empty strings)
        /// </summary>
        private static Gen<string> ValidSurnameGen()
        {
            return Gen.Choose(1, 30)
                .SelectMany(length => Gen.ArrayOf(length, Gen.Elements("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ".ToCharArray())))
                .Where(chars => chars.Length > 0)
                .Select(chars => new string(chars).Trim())
                .Where(surname => !string.IsNullOrWhiteSpace(surname));
        }

        /// <summary>
        /// Custom generator for volunteer dictionaries
        /// </summary>
        private static Gen<Dictionary<string, string>> VolunteerDictionaryGen()
        {
            var volunteerGen = from surname in ValidSurnameGen()
                              from email in ValidEmailGen()
                              select (surname, email);

            return Gen.Choose(0, 20)
                .SelectMany(size => Gen.ListOf(size, volunteerGen))
                .Select(volunteers => volunteers
                    .GroupBy(v => v.surname)
                    .Select(g => g.First())
                    .ToDictionary(v => v.surname, v => v.email));
        }

        // Feature: volunteer-email-notifications, Property 1: Volunteer File Persistence Round Trip
        /// <summary>
        /// Property 1: Volunteer File Persistence Round Trip
        /// For any valid volunteer file data (dictionary of surname to email mappings),
        /// saving the data to a file and then loading it back should produce an equivalent
        /// data structure.
        /// **Validates: Requirements 1.4, 7.1**
        /// </summary>
        [Test]
        public void Property_VolunteerFilePersistenceRoundTrip()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> volunteers) =>
                {
                    var tempFile = Path.GetTempFileName();
                    try
                    {
                        // Act - Save volunteers to file
                        _volunteerManager.SaveVolunteers(tempFile, volunteers);

                        // Act - Load volunteers from file
                        var loaded = _volunteerManager.LoadVolunteers(tempFile);

                        // Assert - Verify equivalence
                        if (volunteers.Count != loaded.Count)
                        {
                            return false.Label($"Count mismatch: expected {volunteers.Count}, got {loaded.Count}");
                        }

                        foreach (var kvp in volunteers)
                        {
                            if (!loaded.ContainsKey(kvp.Key))
                            {
                                return false.Label($"Missing surname in loaded data: {kvp.Key}");
                            }

                            if (loaded[kvp.Key] != kvp.Value)
                            {
                                return false.Label($"Email mismatch for {kvp.Key}: expected {kvp.Value}, got {loaded[kvp.Key]}");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Round trip failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 19: Add Contact Increases Count
        /// <summary>
        /// Property 19: Add Contact Increases Count
        /// For any valid volunteer contact (non-empty surname and valid email),
        /// adding the contact should increase the volunteer count by one and
        /// the addition should be persisted to storage.
        /// **Validates: Requirements 8.4, 8.5**
        /// </summary>
        [Test]
        public void Property_AddContactIncreasesCount()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                Arb.From(ValidSurnameGen()),
                Arb.From(ValidEmailGen()),
                (Dictionary<string, string> initialVolunteers, string newSurname, string newEmail) =>
                {
                    // Skip if the surname already exists in the initial dictionary
                    if (initialVolunteers.ContainsKey(newSurname))
                    {
                        return true.ToProperty().Label("Skipped: surname already exists");
                    }

                    var tempFile = Path.GetTempFileName();
                    try
                    {
                        // Arrange - Save initial volunteers to file
                        _volunteerManager.SaveVolunteers(tempFile, initialVolunteers);
                        var initialCount = initialVolunteers.Count;

                        // Act - Add new volunteer
                        var volunteers = new Dictionary<string, string>(initialVolunteers);
                        _volunteerManager.AddVolunteer(newSurname, newEmail, volunteers);

                        // Assert - Verify count increased by 1
                        if (volunteers.Count != initialCount + 1)
                        {
                            return false.Label($"Count did not increase by 1: expected {initialCount + 1}, got {volunteers.Count}");
                        }

                        // Assert - Verify the volunteer was added correctly
                        if (!volunteers.ContainsKey(newSurname))
                        {
                            return false.Label($"Volunteer with surname '{newSurname}' was not added");
                        }

                        if (volunteers[newSurname] != newEmail)
                        {
                            return false.Label($"Email mismatch: expected '{newEmail}', got '{volunteers[newSurname]}'");
                        }

                        // Act - Persist to storage
                        _volunteerManager.SaveVolunteers(tempFile, volunteers);

                        // Act - Load from storage
                        var loadedVolunteers = _volunteerManager.LoadVolunteers(tempFile);

                        // Assert - Verify persistence
                        if (loadedVolunteers.Count != initialCount + 1)
                        {
                            return false.Label($"Persisted count incorrect: expected {initialCount + 1}, got {loadedVolunteers.Count}");
                        }

                        if (!loadedVolunteers.ContainsKey(newSurname))
                        {
                            return false.Label($"Volunteer '{newSurname}' was not persisted");
                        }

                        if (loadedVolunteers[newSurname] != newEmail)
                        {
                            return false.Label($"Persisted email mismatch: expected '{newEmail}', got '{loadedVolunteers[newSurname]}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Add volunteer failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 21: Delete Contact Removes Entry
        /// <summary>
        /// Property 21: Delete Contact Removes Entry
        /// For any volunteer contact in the list, clicking the delete button for that contact
        /// should remove it from the list, update the file, and persist the change to storage.
        /// **Validates: Requirements 8.10, 8.11**
        /// </summary>
        [Test]
        public void Property_DeleteContactRemovesEntry()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> volunteers) =>
                {
                    // Skip if the dictionary is empty (nothing to delete)
                    if (volunteers.Count == 0)
                    {
                        return true.ToProperty().Label("Skipped: empty volunteer dictionary");
                    }

                    var tempFile = Path.GetTempFileName();
                    try
                    {
                        // Arrange - Save initial volunteers to file
                        _volunteerManager.SaveVolunteers(tempFile, volunteers);
                        var initialCount = volunteers.Count;

                        // Pick a random volunteer to delete
                        var volunteersArray = volunteers.ToArray();
                        var randomIndex = new System.Random().Next(volunteersArray.Length);
                        var volunteerToDelete = volunteersArray[randomIndex];
                        var surnameToDelete = volunteerToDelete.Key;
                        var emailToDelete = volunteerToDelete.Value;

                        // Act - Remove the volunteer
                        var updatedVolunteers = new Dictionary<string, string>(volunteers);
                        _volunteerManager.RemoveVolunteer(surnameToDelete, updatedVolunteers);

                        // Assert - Verify count decreased by 1
                        if (updatedVolunteers.Count != initialCount - 1)
                        {
                            return false.Label($"Count did not decrease by 1: expected {initialCount - 1}, got {updatedVolunteers.Count}");
                        }

                        // Assert - Verify the volunteer was removed
                        if (updatedVolunteers.ContainsKey(surnameToDelete))
                        {
                            return false.Label($"Volunteer with surname '{surnameToDelete}' was not removed");
                        }

                        // Assert - Verify all other volunteers remain
                        foreach (var kvp in volunteers)
                        {
                            if (kvp.Key != surnameToDelete)
                            {
                                if (!updatedVolunteers.ContainsKey(kvp.Key))
                                {
                                    return false.Label($"Volunteer '{kvp.Key}' was incorrectly removed");
                                }

                                if (updatedVolunteers[kvp.Key] != kvp.Value)
                                {
                                    return false.Label($"Email for '{kvp.Key}' was modified: expected '{kvp.Value}', got '{updatedVolunteers[kvp.Key]}'");
                                }
                            }
                        }

                        // Act - Persist to storage
                        _volunteerManager.SaveVolunteers(tempFile, updatedVolunteers);

                        // Act - Load from storage
                        var loadedVolunteers = _volunteerManager.LoadVolunteers(tempFile);

                        // Assert - Verify persistence
                        if (loadedVolunteers.Count != initialCount - 1)
                        {
                            return false.Label($"Persisted count incorrect: expected {initialCount - 1}, got {loadedVolunteers.Count}");
                        }

                        if (loadedVolunteers.ContainsKey(surnameToDelete))
                        {
                            return false.Label($"Deleted volunteer '{surnameToDelete}' was still persisted");
                        }

                        // Assert - Verify all other volunteers were persisted correctly
                        foreach (var kvp in volunteers)
                        {
                            if (kvp.Key != surnameToDelete)
                            {
                                if (!loadedVolunteers.ContainsKey(kvp.Key))
                                {
                                    return false.Label($"Volunteer '{kvp.Key}' was not persisted");
                                }

                                if (loadedVolunteers[kvp.Key] != kvp.Value)
                                {
                                    return false.Label($"Persisted email for '{kvp.Key}' incorrect: expected '{kvp.Value}', got '{loadedVolunteers[kvp.Key]}'");
                                }
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Remove volunteer failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            ).Check(config);
        }
    }
}
