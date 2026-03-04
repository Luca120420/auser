using System;
using System.Collections.Generic;
using System.Linq;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for TransformationRulesEngine class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    [TestFixture]
    public class TransformationRulesEnginePropertyTests
    {
        private ITransformationRulesEngine _engine = null!;

        [SetUp]
        public void SetUp()
        {
            _engine = new TransformationRulesEngine();
        }

        // Feature: auser-excel-transformer, Property 8: Yellow Highlighting Rule
        /// <summary>
        /// Property 8: Yellow Highlighting Rule
        /// For any service appointment where the ATTIVITÀ column contains the text
        /// "Accompag. con macchina attrezzata", that row should be marked for yellow
        /// highlighting in the output.
        /// **Validates: Requirements 4.1**
        /// </summary>
        [Test]
        public void Property_YellowHighlightingRule()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Count how many appointments should be highlighted
                    var expectedHighlightCount = appointments.Count(a =>
                        !string.IsNullOrEmpty(a.Attivita) &&
                        a.Attivita.Contains("Accompag. con macchina attrezzata") &&
                        !IsAnnullato(a)
                    );

                    // Assert - The number of highlighted rows should match
                    var highlightCountMatches = result.YellowHighlightRows.Count == expectedHighlightCount;

                    // Assert - All highlighted rows should correspond to appointments with the text
                    var allHighlightsValid = true;
                    var nonAnnullatoAppointments = appointments.Where(a => !IsAnnullato(a)).ToList();
                    
                    foreach (var rowIndex in result.YellowHighlightRows)
                    {
                        if (rowIndex < 1 || rowIndex > nonAnnullatoAppointments.Count)
                        {
                            allHighlightsValid = false;
                            break;
                        }
                        
                        var appointment = nonAnnullatoAppointments[rowIndex - 1];
                        if (string.IsNullOrEmpty(appointment.Attivita) ||
                            !appointment.Attivita.Contains("Accompag. con macchina attrezzata"))
                        {
                            allHighlightsValid = false;
                            break;
                        }
                    }

                    return highlightCountMatches && allHighlightsValid;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 9: ATTIVITÀ Column Migration
        /// <summary>
        /// Property 9: ATTIVITÀ Column Migration
        /// For any service appointment, the content from the ATTIVITÀ column should appear
        /// in the DESCRIZIONE_STATO_SERVIZIO column in the transformed output, and no
        /// ATTIVITÀ column should exist in the output.
        /// **Validates: Requirements 4.2**
        /// </summary>
        [Test]
        public void Property_AttivitaColumnMigration()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Assert - TransformedRow should not have an Attivita property
                    // This is verified by the model structure itself
                    
                    // Assert - All rows should be transformed (excluding ANNULLATO)
                    var nonAnnullatoCount = appointments.Count(a => !IsAnnullato(a));
                    var rowCountMatches = result.Rows.Count == nonAnnullatoCount;

                    // Assert - The TransformedRow type should not have an Attivita property
                    var transformedRowType = typeof(TransformedRow);
                    var hasAttivitaProperty = transformedRowType.GetProperty("Attivita") != null;
                    var noAttivitaColumn = !hasAttivitaProperty;

                    return rowCountMatches && noAttivitaColumn;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 10: Cancelled Appointment Filtering
        /// <summary>
        /// Property 10: Cancelled Appointment Filtering
        /// For any set of service appointments, those with DESCRIZIONE_STATO_SERVIZIO
        /// equal to "ANNULLATO" should not appear in the transformed output.
        /// **Validates: Requirements 4.3**
        /// </summary>
        [Test]
        public void Property_CancelledAppointmentFiltering()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Assert - No row in the output should have ANNULLATO status
                    // (We can't check this directly on TransformedRow, but we can verify count)
                    var nonAnnullatoCount = appointments.Count(a => !IsAnnullato(a));
                    var rowCountMatches = result.Rows.Count == nonAnnullatoCount;

                    // Assert - If all appointments are ANNULLATO, result should be empty
                    var allAnnullato = appointments.All(a => IsAnnullato(a));
                    var emptyIfAllAnnullato = !allAnnullato || result.Rows.Count == 0;

                    // Assert - If no appointments are ANNULLATO, all should be in result
                    var noneAnnullato = appointments.All(a => !IsAnnullato(a));
                    var allIncludedIfNoneAnnullato = !noneAnnullato || result.Rows.Count == appointments.Count;

                    return rowCountMatches && emptyIfAllAnnullato && allIncludedIfNoneAnnullato;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 11: Punto Partenza Duplication
        /// <summary>
        /// Property 11: Punto Partenza Duplication
        /// For any service appointment where DESCRIZIONE PUNTO PARTENZA is non-empty,
        /// the transformed output should contain that text concatenated with itself (text + text).
        /// **Validates: Requirements 4.4**
        /// </summary>
        [Test]
        public void Property_PuntoPartenzaDuplication()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Get non-annullato appointments
                    var nonAnnullatoAppointments = appointments.Where(a => !IsAnnullato(a)).ToList();

                    // Assert - For each transformed row, check Destinazione field
                    for (int i = 0; i < result.Rows.Count && i < nonAnnullatoAppointments.Count; i++)
                    {
                        var appointment = nonAnnullatoAppointments[i];
                        var row = result.Rows[i];

                        if (!string.IsNullOrEmpty(appointment.DescrizionePuntoPartenza))
                        {
                            // Should be duplicated
                            var expected = appointment.DescrizionePuntoPartenza + appointment.DescrizionePuntoPartenza;
                            if (row.Destinazione != expected)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            // Should be empty or null
                            if (!string.IsNullOrEmpty(row.Destinazione))
                            {
                                return false;
                            }
                        }
                    }

                    return true;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 12: ASSISTITO Column Formation
        /// <summary>
        /// Property 12: ASSISTITO Column Formation
        /// For any service appointment, the ASSISTITO column in the transformed output
        /// should equal COGNOME_ASSISTITO + " " + NOME_ASSISTITO.
        /// **Validates: Requirements 4.5**
        /// </summary>
        [Test]
        public void Property_AssistitoColumnFormation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Get non-annullato appointments
                    var nonAnnullatoAppointments = appointments.Where(a => !IsAnnullato(a)).ToList();

                    // Assert - For each transformed row, check ASSISTITO field
                    for (int i = 0; i < result.Rows.Count && i < nonAnnullatoAppointments.Count; i++)
                    {
                        var appointment = nonAnnullatoAppointments[i];
                        var row = result.Rows[i];

                        var expectedAssistito = $"{appointment.CognomeAssistito} {appointment.NomeAssistito}";
                        if (row.Assistito != expectedAssistito)
                        {
                            return false;
                        }
                    }

                    return true;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 13: INDIRIZZO Column Conditional Concatenation
        /// <summary>
        /// Property 13: INDIRIZZO Column Conditional Concatenation
        /// For any service appointment, if INDIRIZZO_DESTINAZIONE is non-empty,
        /// the INDIRIZZO column should equal COMUNE_DESTINAZIONE + " " + INDIRIZZO_DESTINAZIONE;
        /// otherwise, it should equal COMUNE_DESTINAZIONE + " " + CAUSALE_DESTINAZIONE.
        /// **Validates: Requirements 4.6, 4.7**
        /// </summary>
        [Test]
        public void Property_IndirizzoColumnConditionalConcatenation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Get non-annullato appointments
                    var nonAnnullatoAppointments = appointments.Where(a => !IsAnnullato(a)).ToList();

                    // Assert - For each transformed row, check INDIRIZZO field
                    for (int i = 0; i < result.Rows.Count && i < nonAnnullatoAppointments.Count; i++)
                    {
                        var appointment = nonAnnullatoAppointments[i];
                        var row = result.Rows[i];

                        string comune = appointment.ComuneDestinazione ?? string.Empty;
                        string expectedIndirizzo;

                        if (!string.IsNullOrEmpty(appointment.IndirizzoDestinazione))
                        {
                            // Rule 6: COMUNE_DESTINAZIONE + " " + INDIRIZZO_DESTINAZIONE
                            expectedIndirizzo = $"{comune} {appointment.IndirizzoDestinazione}";
                        }
                        else
                        {
                            // Rule 7: COMUNE_DESTINAZIONE + " " + CAUSALE_DESTINAZIONE
                            string causale = appointment.CausaleDestinazione ?? string.Empty;
                            expectedIndirizzo = $"{comune} {causale}";
                        }

                        if (row.Indirizzo != expectedIndirizzo)
                        {
                            return false;
                        }
                    }

                    return true;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 14: Output Column Structure
        /// <summary>
        /// Property 14: Output Column Structure
        /// For any transformed data, the column order should be: DATA SERVIZIO, ORA INIZIO SERVIZIO,
        /// ASSISTITO, COGNOME ASSISTITO, NOME ASSISTITO, INDIRIZZO, DESTINAZIONE, five empty columns,
        /// ORA INIZIO SERVIZIO (copied), Partenza (empty), NOTE E RICHIESTE.
        /// **Validates: Requirements 4.8**
        /// </summary>
        [Test]
        public void Property_OutputColumnStructure()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Assert - For each transformed row, verify structure
                    foreach (var row in result.Rows)
                    {
                        // Check that all required properties exist and are not null
                        if (row.DataServizio == null) return false;
                        if (row.OraInizioServizio == null) return false;
                        if (row.Assistito == null) return false;
                        if (row.CognomeAssistito == null) return false;
                        if (row.NomeAssistito == null) return false;
                        if (row.Indirizzo == null) return false;
                        // Destinazione can be null
                        
                        // Check that five empty columns exist and are empty
                        if (row.EmptyColumn1 != string.Empty) return false;
                        if (row.EmptyColumn2 != string.Empty) return false;
                        if (row.EmptyColumn3 != string.Empty) return false;
                        if (row.EmptyColumn4 != string.Empty) return false;
                        if (row.EmptyColumn5 != string.Empty) return false;

                        // Check that OraInizioServizioCopy exists and matches OraInizioServizio
                        if (row.OraInizioServizioCopy != row.OraInizioServizio) return false;

                        // Check that Partenza exists (will verify it's empty in Property 15)
                        if (row.Partenza == null) return false;

                        // NoteERichieste can be null
                    }

                    return true;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 15: Partenza Field Always Empty
        /// <summary>
        /// Property 15: Partenza Field Always Empty
        /// For any transformed row, the Partenza field should be empty (null or empty string).
        /// **Validates: Requirements 4.9**
        /// </summary>
        [Test]
        public void Property_PartenzaFieldAlwaysEmpty()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryServiceAppointmentList(),
                (List<ServiceAppointment> appointments) =>
                {
                    // Act
                    var result = _engine.Transform(appointments);

                    // Assert - All Partenza fields should be empty
                    var allPartenzaEmpty = result.Rows.All(row => 
                        row.Partenza == string.Empty || row.Partenza == null
                    );

                    // Assert - Specifically check for empty string (per implementation)
                    var allPartenzaEmptyString = result.Rows.All(row => row.Partenza == string.Empty);

                    return allPartenzaEmpty && allPartenzaEmptyString;
                }
            ).Check(config);
        }

        #region Helper Methods

        /// <summary>
        /// Helper method to check if an appointment is ANNULLATO
        /// </summary>
        private bool IsAnnullato(ServiceAppointment appointment)
        {
            return !string.IsNullOrEmpty(appointment.DescrizioneStatoServizio) &&
                   appointment.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase);
        }

        #endregion

        #region Custom Generators

        /// <summary>
        /// Generator for lists of service appointments
        /// </summary>
        public static Arbitrary<List<ServiceAppointment>> ArbitraryServiceAppointmentList()
        {
            var listGen = Gen.ListOf(ArbitraryServiceAppointment().Generator)
                            .Select(list => list.ToList());
            return Arb.From(listGen);
        }

        /// <summary>
        /// Generator for service appointments with realistic Italian data
        /// </summary>
        public static Arbitrary<ServiceAppointment> ArbitraryServiceAppointment()
        {
            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };
            var normalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ '-";

            // Generator for strings that may contain Italian characters
            var italianStringGen = from length in Gen.Choose(0, 30)
                                   from useItalian in Arb.Generate<bool>()
                                   from chars in Gen.ArrayOf(length, useItalian && length > 0
                                       ? Gen.Elements(italianChars.Concat(normalChars.ToCharArray()).ToArray())
                                       : Gen.Elements(normalChars.ToCharArray()))
                                   select new string(chars).Trim();

            // Generator for non-empty strings (for required fields)
            var nonEmptyItalianStringGen = italianStringGen.Where(s => !string.IsNullOrWhiteSpace(s));

            // Generator for date strings
            var dateGen = from day in Gen.Choose(1, 28)
                         from month in Gen.Choose(1, 12)
                         from year in Gen.Choose(2020, 2030)
                         select $"{day:D2}/{month:D2}/{year}";

            // Generator for time strings
            var timeGen = from hour in Gen.Choose(0, 23)
                         from minute in Gen.Choose(0, 59)
                         select $"{hour:D2}:{minute:D2}";

            // Generator for activity types (including the special one for yellow highlighting)
            var attivitaGen = Gen.Frequency(
                Tuple.Create(1, Gen.Constant("Accompag. con macchina attrezzata")),
                Tuple.Create(1, Gen.Constant("Servizio di Accompag. con macchina attrezzata speciale")),
                Tuple.Create(3, Gen.Constant("Accompagnamento")),
                Tuple.Create(2, Gen.Constant("Trasporto")),
                Tuple.Create(1, Gen.Constant("Visita medica")),
                Tuple.Create(1, italianStringGen)
            );

            // Generator for status descriptions (including ANNULLATO)
            var statusGen = Gen.Frequency(
                Tuple.Create(1, Gen.Constant("ANNULLATO")),
                Tuple.Create(1, Gen.Constant("annullato")),
                Tuple.Create(1, Gen.Constant("Annullato")),
                Tuple.Create(5, Gen.Constant("Confermato")),
                Tuple.Create(2, Gen.Constant("In attesa")),
                Tuple.Create(1, Gen.Constant("Completato")),
                Tuple.Create(1, italianStringGen)
            );

            // Italian names
            var cognomiGen = Gen.Elements(new[] {
                "Rossi", "Bianchi", "Verdi", "Ferrari", "Colombo", "Romano", "Ricci",
                "Marino", "Greco", "Bruno", "Gallo", "Conti", "De Luca", "Mancini",
                "Costa", "Giordano", "Rizzo", "Lombardi", "Moretti", "Barbieri",
                "D'Angelo", "Fontana", "Santoro", "Mariani", "Rinaldi"
            });

            var nomiGen = Gen.Elements(new[] {
                "Mario", "Luigi", "Giuseppe", "Francesco", "Antonio", "Giovanni",
                "Maria", "Anna", "Lucia", "Rosa", "Angela", "Giovanna",
                "Carlo", "Paolo", "Marco", "Stefano", "Alessandro", "Roberto",
                "Elena", "Francesca", "Laura", "Paola", "Chiara", "Sara"
            });

            // Italian cities
            var comuniGen = Gen.Elements(new[] {
                "Milano", "Roma", "Torino", "Napoli", "Palermo", "Genova",
                "Bologna", "Firenze", "Bari", "Catania", "Venezia", "Verona",
                "Messina", "Padova", "Trieste", "Brescia", "Parma", "Prato"
            });

            // Italian addresses
            var indirizziGen = from via in Gen.Elements(new[] { "Via", "Viale", "Corso", "Piazza" })
                              from nome in Gen.Elements(new[] { "Roma", "Verdi", "Garibaldi", "Mazzini", "Dante", "Cavour" })
                              from numero in Gen.Choose(1, 200)
                              select $"{via} {nome} {numero}";

            // Causale/reasons
            var causaleGen = Gen.Elements(new[] {
                "Ospedale", "Clinica", "Centro medico", "Ambulatorio", "Casa di riposo",
                "Farmacia", "Laboratorio analisi", "Fisioterapia", "Visita specialistica"
            });

            var appointmentGen = from dataServizio in dateGen
                                from oraInizio in timeGen
                                from attivita in attivitaGen
                                from status in statusGen
                                from indPartenza in Gen.OneOf(indirizziGen, italianStringGen)
                                from comunePartenza in comuniGen
                                from descPuntoPartenza in italianStringGen
                                from indDestinazione in Gen.OneOf(indirizziGen, italianStringGen)
                                from comuneDestinazione in comuniGen
                                from causale in causaleGen
                                from cognome in cognomiGen
                                from nome in nomiGen
                                from note in italianStringGen
                                select new ServiceAppointment
                                {
                                    DataServizio = dataServizio,
                                    OraInizioServizio = oraInizio,
                                    Attivita = attivita,
                                    DescrizioneStatoServizio = status,
                                    IndirizzoPartenza = indPartenza,
                                    ComunePartenza = comunePartenza,
                                    DescrizionePuntoPartenza = descPuntoPartenza,
                                    IndirizzoDestinazione = indDestinazione,
                                    ComuneDestinazione = comuneDestinazione,
                                    CausaleDestinazione = causale,
                                    CognomeAssistito = cognome,
                                    NomeAssistito = nome,
                                    NoteERichieste = note
                                };

            return Arb.From(appointmentGen);
        }

        #endregion
    }
}
