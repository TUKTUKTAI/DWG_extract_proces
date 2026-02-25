import java.awt.Point;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class P22_0002_Main {

	static LinkedHashMap<String, ArrayList<String[]>> objecten = new LinkedHashMap<String, ArrayList<String[]>>();
	static LinkedHashMap<String, ArrayList<Object[]>> verzameling = new LinkedHashMap<String, ArrayList<Object[]>>();
	static LinkedHashMap<String, ArrayList<String>> verzameling2 = new LinkedHashMap<String, ArrayList<String>>();
	static LinkedHashMap<String, String[]> klokstanden = new LinkedHashMap<String, String[]>();
	static LinkedHashMap<String, ArrayList<String[]>> objectenTotaal = new LinkedHashMap<String, ArrayList<String[]>>();
	static String[] results;
	static String BPScodering;
	static File doelLocatie;
	static File[] inputBestandenVoorRun;

	static int SpecsheetRow = 1;
	static int ObjectRow = 1;
	static int ArmRow = 1;
	static int WerkzaamhedenRow = 1;
	static int AsBuiltRow = 1;
	static int bordSpec = 0;
	static boolean verschovenBord = false;

	static boolean firstPoly = true;
	static FileInputStream fsIP;
	static BufferedInputStream bIP;
	static Workbook workbook;
	static Sheet sheet;
	static Row row;
	static Cell cell;
	static CellStyle cs;

	public static void main(String[] args) throws FileNotFoundException {
		File[] inputBestanden = bepaalInputBestanden(args);
		if (inputBestanden == null || inputBestanden.length == 0) {
			System.out.println("Geen extractbestanden gevonden om te verwerken.");
			return;
		}

		int success = 0;
		for (int i = 0; i < inputBestanden.length; i++) {
			File inputBestand = inputBestanden[i];
			if (inputBestand == null || inputBestand.isFile() == false) {
				continue;
			}
			System.out.println("Verwerken (" + (i + 1) + "/" + inputBestanden.length + "): " + inputBestand.getName());
			try {
				processExtractBestand(inputBestand);
				success++;
			} catch (Exception e) {
				System.out.println("Fout bij " + inputBestand.getName() + ": " + e.getMessage());
				logError(inputBestand, e);
			} finally {
				inputBestandenVoorRun = null;
			}
		}

		System.out.println("Klaar. Succesvol verwerkt: " + success + " van " + inputBestanden.length);
	}

	private static File[] bepaalInputBestanden(String[] args) {
		File inputPad;
		if (args != null && args.length > 0 && args[0] != null && args[0].isBlank() == false) {
			inputPad = new File(args[0]);
		} else {
			inputPad = new File("Bron");
		}

		if (inputPad.exists() == false) {
			System.out.println("Inputpad bestaat niet: " + inputPad.getPath());
			return new File[0];
		}

		if (inputPad.isFile()) {
			return new File[] { inputPad };
		}

		File[] fileList = inputPad.listFiles();
		if (fileList == null) {
			return new File[0];
		}

		ArrayList<File> resultaten = new ArrayList<File>();
		for (int i = 0; i < fileList.length; i++) {
			File f = fileList[i];
			if (f != null && f.isFile()) {
				String naam = f.getName().toLowerCase();
				if (naam.endsWith(".xls") || naam.endsWith(".xlsx")) {
				resultaten.add(f);
				}
			}
		}

		return resultaten.toArray(new File[0]);
	}

	private static void processExtractBestand(File inputBestand) throws IOException {
		resetRunState();
		inputBestandenVoorRun = new File[] { inputBestand };

		File printExcelBlanco = new File("assets" + File.separator + "blanco excels" + File.separator + "blanco resultaat.xlsx");
		doelLocatie = maakDoelBestand(inputBestand);

		Path copied = Paths.get(doelLocatie.getPath());
		Path originalPath = printExcelBlanco.toPath();
		Files.copy(originalPath, copied);

		FileOutputStream output_file = null;
		try {
			fsIP = new FileInputStream(doelLocatie);
			bIP = new BufferedInputStream(fsIP);
			workbook = new XSSFWorkbook(bIP);
			cs = workbook.createCellStyle();
			cs.setWrapText(true);

			readExtracts();
			printObjecten();

			closeQuietly(bIP);
			bIP = null;
			closeQuietly(fsIP);
			fsIP = null;

			output_file = new FileOutputStream(doelLocatie);
			workbook.write(output_file);
		} finally {
			closeQuietly(output_file);
			closeQuietly(workbook);
			closeQuietly(bIP);
			closeQuietly(fsIP);
			workbook = null;
			sheet = null;
			row = null;
			cell = null;
			cs = null;
			bIP = null;
			fsIP = null;
			inputBestandenVoorRun = null;
		}
	}

	private static File maakDoelBestand(File inputBestand) {
		File doelMap = new File("Doel");
		if (doelMap.exists() == false) {
			doelMap.mkdirs();
		}

		String naam = "extract";
		if (inputBestand != null) {
			naam = inputBestand.getName();
			int idx = naam.lastIndexOf('.');
			if (idx > 0) {
				naam = naam.substring(0, idx);
			}
		}
		naam = naam.replaceAll("[\\\\/:*?\"<>|]", "_");

		File output = new File(doelMap, "Extraheren resultaat " + naam + ".xlsx");
		int counter = 2;
		while (output.exists()) {
			output = new File(doelMap, "Extraheren resultaat " + naam + " (" + counter + ").xlsx");
			counter++;
		}
		return output;
	}

	private static void resetRunState() {
		objecten.clear();
		verzameling.clear();
		verzameling2.clear();
		klokstanden.clear();
		objectenTotaal.clear();
		results = null;
		BPScodering = "";

		SpecsheetRow = 1;
		ObjectRow = 1;
		ArmRow = 1;
		WerkzaamhedenRow = 1;
		AsBuiltRow = 1;
		bordSpec = 0;
		verschovenBord = false;
		firstPoly = true;
	}

	private static void logError(File inputBestand, Exception e) {
		try {
			File doelMap = new File("Doel");
			if (doelMap.exists() == false) {
				doelMap.mkdirs();
			}
			Path errorBestand = Paths.get(doelMap.getPath(), "errors.csv");
			String tijd = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
			String pad = inputBestand == null ? "" : inputBestand.getPath();
			String melding = e == null ? "" : (e.getClass().getSimpleName() + ": " + String.valueOf(e.getMessage()));
			if (e != null && e.getCause() != null && e.getCause().getMessage() != null) {
				melding = melding + " | cause=" + e.getCause().getClass().getSimpleName() + ": " + e.getCause().getMessage();
			}
			melding = melding.replace("\r", " ").replace("\n", " ").replace(";", ",");

			String regel = tijd + ";" + pad.replace(";", ",") + ";" + melding + System.lineSeparator();
			Files.write(errorBestand, regel.getBytes(StandardCharsets.UTF_8), StandardOpenOption.CREATE,
					StandardOpenOption.APPEND);
		} catch (IOException io) {
			System.out.println("Kan errors.csv niet schrijven: " + io.getMessage());
		}
	}

	private static void closeQuietly(AutoCloseable resource) {
		if (resource != null) {
			try {
				resource.close();
			} catch (Exception e) {
			}
		}
	}

	
	private static void verwerkExtracts() {
		Iterator it = objecten.entrySet().iterator();

		int cnt = 0;
		while (it.hasNext()) {
			
			results = new String[32];
			BPScodering="";
			verzameling.clear();
			verzameling2.clear();
			klokstanden.clear();
			System.gc();
			HashMap.Entry<String, ArrayList<String[]>> pair = (HashMap.Entry) it.next();
			ArrayList<String[]> bord = pair.getValue();
			cnt++;
			
			// System.out.println(cnt); //pair.getKey()
			//System.out.println(pair.getKey());
			// [0] file name
			// [1] file location
			// [2] name
			// [3] x pos
			// [4] y pos
			// [5] z pos
			// [6] x start
			// [7] y start
			// [8] z start
			// [9] x end
			// [10] y end
			// [11] z end
			// [12] DATUM
			// [14] GETEKEND
			// [15] GEZIEN
			// [16] REVISIE
			// [17] SOORT
			// [18] SOORT_HANDW
			// [19] value
			// [20] layer
			// [21] length
			// [22] rotation
			// [23] date modified
			// [24] CAMTAG

			bordSpec = 4;
			boolean borddata = false;
			firstPoly = true;
			verschovenBord = false;
			double xSchuiving = 0;
			
			//System.out.println(bord.get(0)[1]);
			boolean geclassificeerd = false;
			for (int i = 0; i < bord.size(); i++) {
				try {
				String[] temp = bord.get(i);
				if (temp[2].startsWith("SPEC01")||temp[2].startsWith("SPEC FIETS")) {
					bordSpec = 1; // standaard
					xSchuiving = Double.valueOf(temp[3]);
					
					geclassificeerd = true;
					break;
				} else if (temp[2].startsWith("SPEC02")) {
					bordSpec = 2;
					xSchuiving = Double.valueOf(temp[3]);
					
					geclassificeerd = true;
					break;
				}  else if (temp[2].startsWith("SPEC03")) {
					geclassificeerd = true;
					bordSpec = 3; // paddestoel
					xSchuiving = Double.valueOf(temp[3]);
					
					break;
				} else if (temp[2].equals("FRSPEC")) {
					bordSpec = -1;
					geclassificeerd = true;
					xSchuiving = Double.valueOf(temp[3]);
					
					break;
				}else if (temp[2].equals("FLSPECI")) {
					bordSpec = -2;
					geclassificeerd = true;
					xSchuiving = Double.valueOf(temp[3]);
					
					break;
				}
				else if (temp[2].startsWith("SPEC")){
					System.out.println("nieuw bordtype "+temp[2]);
				}
				}catch(Exception e) {
					System.out.println("test bordspec: "+e.getMessage());
				}
			}
			if(geclassificeerd == false && (bord.get(0)[0].contains("T.dwg")||bord.get(0)[0].contains("T.DWG"))) {
				bordSpec = 5;
				try {
				results[6] = bord.get(0)[0].replaceAll("T.dwg", "").replaceAll("T.DWG", "");
				results[7] = bord.get(0)[0].replaceAll("T.dwg", "").replaceAll("T.DWG", "");
				for (int i = 0; i < bord.size(); i++) {
					String[] temp = bord.get(i);
					if(temp[2].equals("BORDDATA")||temp[2].equals("BORDDATA1")) {
						borddata = true;
						break;
					}
				}
				}catch(Exception e) {
					System.out.println("test bordspec 51: "+e.getMessage());
				}
			}
			else if(geclassificeerd == false) {
				
				//System.out.println("geen bordtype: "+bord.get(0)[0]);
				
				bordSpec = 1; // standaard
				try {
				for (int i = 0; i < bord.size(); i++) {
					String[] temp = bord.get(i);
					if(temp[2].equals("Text")&&temp[19].contentEquals("Opmerkingen")) {
						bordSpec = 2; 
						break;
					}
				}
				xSchuiving = 0;
				
				}catch(Exception e) {
					System.out.println("test bordspec 11: "+e.getMessage());
				}
			}
			for (int i = 0; i < bord.size(); i++) {
				
				try {
				String[] temp = bord.get(i);
				if (temp[2].equals("Text") || (!temp[26].equals("-1") && !temp[26].isBlank()) || !temp[24].isBlank()) {

					if ((!temp[26].equals("-1") && !temp[26].isBlank())) {
						temp[19] = temp[26];
					}
					if (temp[3] != null && temp[4] != null && !temp[3].isEmpty() && !temp[4].isEmpty()) {
						double X = Double.valueOf(temp[3]);
						double Y = Double.valueOf(temp[4]);

						if ((X-xSchuiving > 21 && X-xSchuiving < 47 && Y > 0 && Y < 28)
								||(X-xSchuiving > 21*1000 && X-xSchuiving < 47*1000 && Y > 0 && Y < 28*1000)) {// Art.code plaatsing
							// System.out.println(temp[19]);
							if (temp[19].startsWith("  ")) {
								verschovenBord = true;
								break;
								// voegToe(X, Y, temp[19].strip(), "aant.");
							}
						}
					}
				}
				}catch(Exception e) {
					System.out.println("test bordspec xx: "+e.getMessage());
				}

			}
			
			
			//System.out.println(bordSpec);
			String stand2 ="";
			for (int i = 0; i < bord.size(); i++) {
				String[] temp = bord.get(i);
				results[0] = temp[1] + "\\" + temp[0];
				results[1] = temp[0];

				if(temp[2].equals("overbetuwe")) {
					results[15]= "RIJSSEN";
				}
				
				if(bordSpec == 3 && temp[2].equals("PADRI")) {
					//paddestoel					
					try {
					String[] tempArray = new String[3];
					
					int tmp = Integer.valueOf(temp[22].split("d")[0].replaceAll("g",""));
					
					tempArray[0] = String.valueOf(tmp);
					
					klokstanden.put("1", tempArray); // paddestoel heeft maar 1 bord
					}catch(Exception e) {
						try {
							String[] tempArray = new String[3];
							
							double tmp2 = Double.valueOf(temp[22].split("d")[0].replaceAll("g",""));
							int tmp = (int) tmp2;
							tempArray[0] = String.valueOf(tmp);
							
							klokstanden.put("1", tempArray); // paddestoel heeft maar 1 bord
						}catch(Exception e2) {
						System.out.println("klokstanden error paddestoel: "+e2.getMessage());
						}
					}
				}
				
				if (temp[2].equals("Line") && (bordSpec==1 || bordSpec==0)) { //klokstanden
					try {
					// mogelijke klokstand
					double X = Double.valueOf(temp[6]); // x coordinaat van bornummer
					
					if ((X-xSchuiving == 40.85 && temp[7].equals("225.5000"))
							||(X-xSchuiving == 40850.0 && temp[7].equals("225500.00"))
							||(X-xSchuiving == 1.6083 && temp[7].equals("8.8780"))
							||(X-xSchuiving == 1608.3 && temp[7].equals("8878.00"))) {// klokstand
						// deze klokstanden komen vanuit het midden
						// double theta = Math.toDegrees(Math.atan2(y - cy, x - cx));
						double theta = Math.toDegrees(Math.atan2(Double.valueOf(temp[10]) - Double.valueOf(temp[7]),
								Double.valueOf(temp[9]) - Double.valueOf(temp[6])));
						;

						String bordnummer = "-1";
						double afstand = 999999;
						double afstandA = 999999;
						double afstandB = 999999;
						double xA = -1;
						double xB = -1;
						double yA = -1;
						double yB = -1;
						for (int z = 0; z < bord.size(); z++) {
							String[] temp2 = bord.get(z);

							if (temp2[2].equals("Text") && !temp2[19].equals("a") && !temp2[19].equals("b")) {

								double x = Double.valueOf(temp2[3]); // x coordinaat van bornummer
								double y = Double.valueOf(temp2[4]); // y coordinaat van bordnummer
								if ((x-xSchuiving > 20 && x-xSchuiving < 62 && y > 208 && y < 242 && !temp[9].isEmpty()
										&& !temp[10].isEmpty())
										||(x-xSchuiving > 20*1000 && x-xSchuiving < 62*1000 && y > 208*1000 && y < 242*1000 && !temp[9].isEmpty()
												&& !temp[10].isEmpty())
										||(x-xSchuiving > 1.18 && x-xSchuiving < 2.1 && y > 8.4 && y < 9.3 && !temp[9].isEmpty()
												&& !temp[10].isEmpty())) {

									double Xend = Double.valueOf(temp[9]); // x coordinaat van eindpunt lijn klokstand
									double Yend = Double.valueOf(temp[10]); // y coordinaat van eindpunt lijn klokstand

									double ac = Math.abs(Yend - y);
									double cb = Math.abs(Xend - x);

									double afstand2 = Math.hypot(ac, cb);
									if (afstand2 < afstand) {
										afstand = afstand2;
										bordnummer = temp2[19];
									}

								}
							} else if (temp2[2].equals("Text") && (temp2[19].equals("a") || temp2[19].equals("b"))) {
								double x = Double.valueOf(temp2[3]); // x coordinaat
								double y = Double.valueOf(temp2[4]); // y coordinaat
								if ((x-xSchuiving > 20 && x-xSchuiving < 62 && y > 208 && y < 242 && !temp[9].isEmpty()
										&& !temp[10].isEmpty())
										||(x-xSchuiving > 20*1000 && x-xSchuiving < 62*1000 && y > 208*1000 && y < 242*1000 && !temp[9].isEmpty()
												&& !temp[10].isEmpty())
										||(x-xSchuiving > 1.18 && x-xSchuiving < 2.1 && y > 8.4 && y < 9.3 && !temp[9].isEmpty()
												&& !temp[10].isEmpty())) {
									double Xend = Double.valueOf(temp[9]); // x coordinaat van eindpunt lijn klokstand
									double Yend = Double.valueOf(temp[10]); // y coordinaat van eindpunt lijn klokstand

									double ac = Math.abs(Yend - y);
									double cb = Math.abs(Xend - x);

									double afstand2 = Math.hypot(ac, cb);
									if (temp2[19].equals("a")) {
										if (afstand2 < afstandA) {
											afstandA = afstand2;

											xA = x;
											yA = y;

										}
									} else if (temp2[19].equals("b")) {
										if (afstand2 < afstandB) {
											afstandB = afstand2;
											xB = x;
											yB = y;
										}
									}

								}

							}
						}
						int angle1 = (int) Math.round(theta);

						if (angle1 >= 0) {
							if (angle1 <= 90) {
								angle1 = 90 - angle1;
							} else {
								angle1 = 90 - angle1;
								angle1 = 360 + angle1;
							}
						} else {
							angle1 = 90 + angle1 * -1;
						}

						if (bordnummer.contains(",")) {
							String[] tempSplit = bordnummer.split(",");
							for (int x = 0; x < tempSplit.length; x++) {
								
								String[] tempArray = new String[3];
								tempArray[0] = String.valueOf(angle1);
								
								klokstanden.put(tempSplit[x], tempArray);
							}
						}else if (bordnummer.contains("-")) {
							String[] tempSplit = bordnummer.split("-");
							for (int x = 0; x < tempSplit.length; x++) {
								
								String[] tempArray = new String[3];
								tempArray[0] = String.valueOf(angle1);
								
								klokstanden.put(tempSplit[x], tempArray);
							}
						} else {
							String[] tempArray = new String[3];
							tempArray[0] = String.valueOf(angle1);
							
							klokstanden.put(bordnummer, tempArray);

							
						}

					} else {
						// een lijn is 2 klokstanden
						double Xstart = Double.valueOf(temp[6]);
						double Ystart = Double.valueOf(temp[7]);
						double Xend = Double.valueOf(temp[9]);
						double Yend = Double.valueOf(temp[10]);
						double length = Double.valueOf(temp[21]);
						double xA1 = -1;
						double xB1 = -1;
						double yA1 = -1;
						double yB1 = -1;
						double xA2 = -1;
						double xB2 = -1;
						double yA2 = -1;
						double yB2 = -1;
						if ((Xstart-xSchuiving > 29 && Xstart-xSchuiving < 52 && Ystart > 214 && Ystart < 237 && Xend-xSchuiving > 29 && Xend-xSchuiving < 52
								&& Yend > 214 && Yend < 237 && length > 20 && length < 24)
								) {

							Double Xmiddle = (Xstart + Xend) / 2;
							Double Ymiddle = (Ystart + Yend) / 2;
							double theta = Math.toDegrees(Math.atan2(Ystart - Ymiddle, Xstart - Xmiddle));
							double theta2 = Math.toDegrees(Math.atan2(Yend - Ymiddle, Xend - Xmiddle));

							String bordnummer = "-1";
							double afstand = 999999;
							double afstandA = 999999;
							double afstandB = 999999;

							String bordnummer2 = "-1";
							double afstand3 = 999999;
							double afstand3A = 999999;
							double afstand3B = 999999;

							for (int z = 0; z < bord.size(); z++) {
								String[] temp2 = bord.get(z);

								if (temp2[2].equals("Text") && !temp2[19].equals("a") && !temp2[19].equals("b")) {

									double x = Double.valueOf(temp2[3]); // x coordinaat van bordnummer
									double y = Double.valueOf(temp2[4]); // y coordinaat van bordnummer
									if (x-xSchuiving > 20 && x-xSchuiving< 62 && y > 208 && y < 242 && !temp[9].isEmpty()
											&& !temp[10].isEmpty()) {

										double ac = Math.abs(Ystart - y);
										double cb = Math.abs(Xstart - x);

										double afstand2 = Math.hypot(ac, cb);
										if (afstand2 < afstand) {
											afstand = afstand2;
											bordnummer = temp2[19];
										}

										double ac2 = Math.abs(Yend - y);
										double cb2 = Math.abs(Xend - x);

										double afstand4 = Math.hypot(ac2, cb2);
										if (afstand4 < afstand3) {
											afstand3 = afstand4;
											bordnummer2 = temp2[19];
										}

									}
								} else if (temp2[2].equals("Text")
										&& (temp2[19].equals("a") || temp2[19].equals("b"))) {
									double x = Double.valueOf(temp2[3]); // x coordinaat van bordnummer
									double y = Double.valueOf(temp2[4]); // y coordinaat van bordnummer
									if (x-xSchuiving > 20 && x-xSchuiving < 62 && y > 208 && y < 242 && !temp[9].isEmpty()
											&& !temp[10].isEmpty()) {
										double ac = Math.abs(Ystart - y);
										double cb = Math.abs(Xstart - x);

										double afstand2 = Math.hypot(ac, cb);
										if (temp2[19].equals("a")) {
											if (afstand2 < afstandA) {
												afstandA = afstand2;

												xA1 = x;
												yA1 = y;
											}
										} else if (temp2[19].equals("b")) {
											if (afstand2 < afstandB) {
												afstandB = afstand2;
												xB1 = x;
												yB1 = y;
											}
										}

										double ac2 = Math.abs(Yend - y);
										double cb2 = Math.abs(Xend - x);

										double afstand4 = Math.hypot(ac2, cb2);
										if (temp2[19].equals("a")) {
											if (afstand4 < afstand3A) {
												afstand3A = afstand4;

												xA2 = x;
												yA2 = y;
											}
										} else if (temp2[19].equals("b")) {
											if (afstand4 < afstand3B) {
												afstand3B = afstand4;
												xB2 = x;
												yB2 = y;
											}
										}

									}
								}
							}
							int angle1 = (int) Math.round(theta);
							int angle2 = (int) Math.round(theta2);
							if (angle1 >= 0) {
								if (angle1 <= 90) {
									angle1 = 90 - angle1;
								} else {
									angle1 = 90 - angle1;
									angle1 = 360 + angle1;
								}
							} else {
								angle1 = 90 + angle1 * -1;
							}

							if (angle2 >= 0) {
								if (angle2 <= 90) {
									angle2 = 90 - angle2;
								} else {
									angle2 = 90 - angle2;
									angle2 = 360 + angle2;
								}
							} else {
								angle2 = 90 + angle2 * -1;
							}
							if (bordnummer.contains(",")) {
								String[] tempSplit = bordnummer.split(",");
								for (int x = 0; x < tempSplit.length; x++) {
									String[] tempArray = new String[3];
									tempArray[0] = String.valueOf(angle1);
									
									klokstanden.put(tempSplit[x], tempArray);
									
								}
							}else if (bordnummer.contains("-")) {
								String[] tempSplit = bordnummer.split("-");
								for (int x = 0; x < tempSplit.length; x++) {
									String[] tempArray = new String[3];
									tempArray[0] = String.valueOf(angle1);
									
									klokstanden.put(tempSplit[x], tempArray);
									
								}
							} else {
								String[] tempArray = new String[3];
								tempArray[0] = String.valueOf(angle1);
								
								klokstanden.put(bordnummer, tempArray);
								
							}

							if (bordnummer2.contains(",")) {
								String[] tempSplit = bordnummer2.split(",");
								for (int x = 0; x < tempSplit.length; x++) {
									String[] tempArray = new String[3];
									tempArray[0] = String.valueOf(angle2);
									
									klokstanden.put(tempSplit[x], tempArray);
									
								}
							}else if (bordnummer2.contains("-")) {
								String[] tempSplit = bordnummer2.split("-");
								for (int x = 0; x < tempSplit.length; x++) {
									String[] tempArray = new String[3];
									tempArray[0] = String.valueOf(angle2);
									
									klokstanden.put(tempSplit[x], tempArray);
									
								}
							} else {
								String[] tempArray = new String[3];
								tempArray[0] = String.valueOf(angle2);
								
								
								klokstanden.put(bordnummer2, tempArray);
								
							}

						}
					}
					}catch(Exception e) {
						System.out.println("klokstanden error: "+e.getMessage());
					}
				} else if (temp[2].startsWith("SPEC")||temp[2].startsWith("FRSPEC")||temp[2].startsWith("FLSPECI") || (bordSpec ==0&& !temp[28].isEmpty())) {
					try {
					if(bordSpec==0 && !temp[28].isEmpty()) {
						if(temp[3] != null && temp[4] != null && !temp[3].isEmpty()&& !temp[4].isEmpty()) {
							double X = Double.valueOf(temp[3]);
							double Y = Double.valueOf(temp[4]);
							
							if((X-xSchuiving>179 && X-xSchuiving<197  && Y>27 && Y<32)||
									(X-xSchuiving>179*1000 && X-xSchuiving<197*1000  && Y>27*1000 && Y<32*1000)) {
								results[2] = temp[28]; // getekend
							}else if((X-xSchuiving>176 && X-xSchuiving<197  && Y>22 && Y<27)||
									(X-xSchuiving>176*1000 && X-xSchuiving<197*1000  && Y>22*1000 && Y<27*1000)) {
								results[3] = temp[28]; // datum
							}else if((X-xSchuiving>176 && X-xSchuiving<197  && Y>17 && Y<22)||
									(X-xSchuiving>176*1000 && X-xSchuiving<197*1000  && Y>17*1000 && Y<22*1000)) {
								results[4] = temp[28]; // gezien
							}else if((X-xSchuiving>176 && X-xSchuiving<197  && Y>12 && Y<17)||
									(X-xSchuiving>176*1000 && X-xSchuiving<197*1000  && Y>12*1000 && Y<17*1000)) {
								results[5] = temp[28]; // revisie
							}
						}
					}else {
					results[2] = temp[14]; // getekend
					
					results[3] = temp[12]; // datum

					results[4] = temp[15]; // gezien

					results[5] = temp[16]; // revisie
					}
					}catch(Exception e) {
						System.out.println("test 6: "+e.getMessage());
					}
				

				}else if(bordSpec == 4 && temp[2].equals("BPS_FILENAME")) {
				
					BPScodering = temp[29]; //CADFIL = bps nummer
				}else if(bordSpec == 4 && temp[2].equals("BORDDATA1")) {
									
					results[7] = temp[31].replaceAll("/", "");
					results[6] = temp[31].replaceAll("/", "");
					voegToe(0, 0, temp[29], "standplaats");
					
					results[18] =temp[32]; //uppercase
					results[19] =temp[33]; //lowercase
					results[20] =temp[34]; //lettertype
					voegToe(0, 0, temp[30].strip(), "Omschrijving");
					
					
					
				}else if(bordSpec == 4 && temp[2].equals("Text")) {
					try {
					if(temp[3] != null && temp[4]!= null) {
						
						double X = Double.valueOf(temp[3]);
						double Y = Double.valueOf(temp[4]);
						
						if(X>100 && X <4400 && Y>0 && Y<4400) {
							voegToe(X, Y, temp[19], "bord_1_a");
						}else if(X>100*1000 && X <4400*1000 && Y>0 && Y<4400*1000) {
							voegToe(X, Y, temp[19], "bord_1_a");
						}
						
					}
					}catch(Exception e) {
						System.out.println("test bordspec 41: "+e.getMessage());
					}
				}else if(bordSpec == 5 && (temp[2].equals("Text")||temp[2].equals("BORDDATA")||temp[2].equals("BORDDATA1")||temp[2].equals("MText"))) {
					
					if(temp[2].equals("Text")) {
						try {
						double X = Double.valueOf(temp[3]);
						double Y = Double.valueOf(temp[4]);
						if(X>0 && Y>0) {
							voegToe(X, Y, temp[19], "bord_1_a");
						}else if(borddata==false && Y<0) {
							voegToe(X, Y, temp[19], "borddata");
						}
						}catch(Exception e) {
							System.out.println("test bordSpec5: "+e.getMessage());
						}
					}else if(temp[2].equals("BORDDATA")||temp[2].equals("BORDDATA1")) {
						
						try {
						
								
						results[18] =temp[32]; //uppercase
						results[19] =temp[33]; //lowercase
						results[20] =temp[34]; //lettertype
						results[30] =temp[35]; //BW
						results[31] =temp[36]; //BH
						voegToe(0, 0, temp[30].strip(), "Omschrijving");
						voegToe(0, 0, temp[31].strip(), "Omschrijving");
						voegToe(0, 0, temp[29], "standplaats");
						}catch(Exception e) {
							System.out.println("test bordSpec5 2: "+e.getMessage());
						}
					}else if(temp[2].equals("MText")&&temp[20].equals("0") && borddata==false) { //geen borddata
						try {
						String tot = temp[19];
						results[6] = bord.get(0)[0].replaceAll("T.dwg", "").replaceAll("T.DWG", "");
						results[7] = bord.get(0)[0].replaceAll("T.dwg", "").replaceAll("T.DWG", "");
						String pattern = "\\\\P";
						if(tot.isEmpty()==false) {
						String[] tot2 = tot.split(pattern);
						voegToe(0, 0, tot2[0].strip(), "standplaats");
						if(tot2.length>1) {
							tot2[1]=tot2[1].replace("BxH(mm):", "");
							String[] tot3 =tot2[1].split("x");
							if(tot3.length>1) {
							results[30] =tot3[0]; //BW
							results[31] =tot3[1]; //BH
							}else {
								results[30] = tot2[1];
							}
						}
						if(tot2.length>2) {
							for (int x=2;x<tot2.length;x++){
								voegToe(0, 0, tot2[x].strip(), "Omschrijving");
							}
						}
						}
						}catch(Exception e) {
							System.out.println("test bordSpec5 3: "+e.getMessage());
						}
					}else if(temp[2].equals("MText") &&temp[3]!= null && temp[3].isEmpty()==false && temp[4] != null&& temp[4].isEmpty()==false){
						try {
						double X = Double.valueOf(temp[3]);
						double Y = Double.valueOf(temp[4]);
						
						if(X>0 && Y>0) {
							if(temp[37].isEmpty()==false) {
							voegToe(X, Y, temp[37].strip(), "omschrijving werkzaamheden");
							}else if(temp[19].isEmpty()==false) {
								voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
							}
						}
						}catch(Exception e) {
							System.out.println("test bordSpec5 4: "+e.getMessage());
						}
					}
					
				}else if ((temp[2].equals("Text") || (!temp[26].equals("-1")&&!temp[26].isBlank())||!temp[24].isBlank())&&temp[3]!=null && temp[4]!= null
						&&!temp[3].isEmpty()&&!temp[4].isEmpty()) {
					
					if((!temp[26].equals("-1")&&!temp[26].isBlank())) {
						temp[19] = temp[26];
					}
					try {
					double X = Double.valueOf(temp[3]);
					double Y = Double.valueOf(temp[4]);
					
					// mogelijke tekst
					if (((X-xSchuiving ==20 ||X-xSchuiving ==40 || X-xSchuiving ==38|| X-xSchuiving ==49) && temp[4].equals("269.0000"))					
							||(X-xSchuiving >=36.5 &&X-xSchuiving <=40.5 &&Y>=268 && Y<=269.5)
							//||(X-xSchuiving ==1.5748 && temp[4].equals("10.5906"))
							||((X-xSchuiving ==40000 || X-xSchuiving ==38000|| X-xSchuiving ==49000) && temp[4].equals("269000.0000"))
							||((X-xSchuiving ==40000 || X-xSchuiving ==38000|| X-xSchuiving ==49000) && temp[4].equals("269000.00"))) { // provincie
						results[14] = temp[19];
					} else if (((X-xSchuiving ==81.0000||X-xSchuiving ==69.0000||X-xSchuiving ==89.0000) && temp[4].equals("269.0000"))
							||(X-xSchuiving >=88 &&X-xSchuiving <=89.5 &&Y>=268 && Y<=269.5)
							//||(X-xSchuiving ==3.5039 && temp[4].equals("10.5906"))
							||((X-xSchuiving ==81000||X-xSchuiving ==89000) && temp[4].equals("269000.0000"))
							||((X-xSchuiving ==81000||X-xSchuiving ==89000) && temp[4].equals("269000.00"))) {// gemeente
						results[15] = temp[19];
					} else if (((X-xSchuiving ==129.5000||X-xSchuiving ==149.5000) && temp[4].equals("269.0000"))
							||(X-xSchuiving >=149 &&X-xSchuiving <=151 &&Y>=268 && Y<=269.5)
							//||(X-xSchuiving ==5.8858 && temp[4].equals("10.5906"))
							||(X-xSchuiving ==149500.0000 && temp[4].equals("269000.0000"))
							||(X-xSchuiving ==149500.0000 && temp[4].equals("269000.00"))) {// soort / nomenclatuur
						results[16] = temp[19];
					} else if ((X-xSchuiving ==177.0000 && temp[4].equals("269.0000"))
							|| (X-xSchuiving ==157.0000 && temp[4].equals("269.0000"))
							|| (X-xSchuiving ==175.0000 && temp[4].equals("269.0000"))
							//|| (X-xSchuiving ==6.9685 && temp[4].equals("10.5906"))
							|| (X-xSchuiving >=177&&X-xSchuiving <=178 &&Y>=268 && Y<=269.5)
							|| (X-xSchuiving ==177000.0000 && temp[4].equals("269000.0000"))
							||(X-xSchuiving ==177000.00 && temp[4].equals("269000.00"))) {// bordnummer
						results[7] = temp[19].replaceAll("/", "");
						results[6] = temp[19].replaceAll("/", "");
					} else if ((X-xSchuiving ==170.0000 && temp[4].equals("3.0000"))||
							(X-xSchuiving ==170000.00 && temp[4].equals("3000.00"))||
							(X-xSchuiving ==170000.00 && temp[4].equals("3000.0000"))) {// bordnummer 2
						results[7] = temp[19].replaceAll("/", "");
						results[6] = temp[19].replaceAll("/", "");
					} else if ((X-xSchuiving ==20.0000 ||X-xSchuiving ==40.0000 ||X-xSchuiving ==49.0000 ||X-xSchuiving ==49000.0000
							||X-xSchuiving ==40000.0000 ||X-xSchuiving ==38.0000|| X-xSchuiving ==38000.0000)
							&& (temp[4].equals("263.0000")||temp[4].equals("263000.0000")||temp[4].equals("263000.00"))
							
							||(X-xSchuiving >=37 &&X-xSchuiving <=41 &&Y>=262 && Y<=263.5)
							//||(X-xSchuiving ==1.5748  &&Y==10.3543)
							) {// wegbeheerder
						
						results[8] = temp[19];
					} else if (((X-xSchuiving ==31.0000||X-xSchuiving ==51.0000) && temp[4].equals("263.0000")) 
							||(X-xSchuiving ==51000.0000 && (temp[4].equals("263000.0000")||temp[4].equals("263000.00")))
							//||(X-xSchuiving ==2.0079  &&Y==10.3543)
							||(X-xSchuiving >=50 &&X-xSchuiving <=52 &&Y>=262 && Y<=263.5)
							) {// wegbeheerder deel 2
						results[9] = temp[19];
					} else if (((X-xSchuiving ==20.0000 ||X-xSchuiving ==40.0000 ||X-xSchuiving ==49.0000 || X-xSchuiving ==38.0000) && temp[4].equals("257.0000"))
							||((X-xSchuiving ==40000.0000 || X-xSchuiving ==38000.0000|| X-xSchuiving ==49000.0000) && (temp[4].equals("257000.0000")||temp[4].equals("257000.00")))						
							||(X-xSchuiving >=37 &&X-xSchuiving <=41 &&Y>=256 && Y<=257.5)
							//||(X-xSchuiving ==1.5748  &&Y==10.1181)
							) {// onderhouder
						results[10] = temp[19];
					} else if (((X-xSchuiving ==31.0000||X-xSchuiving ==51.0000) && temp[4].equals("257.0000"))
							||(X-xSchuiving ==51000.0000 && (temp[4].equals("257000.0000")||temp[4].equals("257000.00")))
							||(X-xSchuiving >=50 &&X-xSchuiving <=52 &&Y>=256 && Y<=257.5)
							//||(X-xSchuiving ==2.0079  &&Y==10.1181)
							) {// onderhouder deel 2
						results[11] = temp[19];
					} else if (((X-xSchuiving ==20.0000 ||X-xSchuiving ==40.0000 ||X-xSchuiving ==49.0000 || X-xSchuiving ==38.0000) && temp[4].equals("251.0000"))
							||((X-xSchuiving ==40000.0000 ||X-xSchuiving ==49000.0000 || X-xSchuiving ==38000.0000) && (temp[4].equals("251000.0000")||temp[4].equals("251000.00")))
							||(X-xSchuiving >=37 &&X-xSchuiving <=41 &&Y>=250 && Y<=251.5)
							//||(X-xSchuiving ==1.5748  &&Y==9.8819)
							) {// off./ord.
						results[12] = temp[19];
					} else if (((X-xSchuiving ==20.0000 ||X-xSchuiving ==40.0000 ||X-xSchuiving ==49.0000 || X-xSchuiving ==38.0000) && temp[4].equals("245.0000"))
							||((X-xSchuiving ==40000.0000||X-xSchuiving ==49000.0000 || X-xSchuiving ==38000.0000) && (temp[4].equals("245000.0000")||temp[4].equals("245000.00")))
							||(X-xSchuiving >=37 &&X-xSchuiving <=41 &&Y>=244 && Y<=245.5)
							//||(X-xSchuiving ==1.5748  &&Y==9.6457)
							) {// PD nr
						results[13] = temp[19];
						
					}else if ((bordSpec ==1||bordSpec==0)&&((X-xSchuiving>20 && X-xSchuiving<197 && Y>193 && Y<208)||(X-xSchuiving>20000 && X-xSchuiving<197000 && Y>193000 && Y<208000))) {// standplaats
						//results[17] = temp[19];
						voegToe(X, Y, temp[19], "standplaats");
						
					}else if ((bordSpec ==2||bordSpec ==-2)&&((X-xSchuiving>20 && X-xSchuiving<160 && Y>148 && Y<167)||(X-xSchuiving>20 && X-xSchuiving<160 && Y>148 && Y<167)||(X-xSchuiving>20000 && X-xSchuiving<160000 && Y>148000 && Y<167000))) {// standplaats
						//results[17] = temp[19];
						if(temp[19].equals("Standplaats")) {
							
						}else {
							voegToe(X, Y, temp[19], "standplaats");
						}
					}else if ((bordSpec ==3)&&((X-xSchuiving>20000 && X-xSchuiving<160000 && Y>148000 && Y<167000)||(X-xSchuiving>20 && X-xSchuiving<160 &&Y>148 && Y<167))) {// standplaats
						//results[17] = temp[19];
						if(temp[19].equals("Standplaats")) {
							
						}else {
							voegToe(X, Y, temp[19], "standplaats");
						}
					}else if ((bordSpec ==-1)&&((X-xSchuiving>20 && X-xSchuiving<130 && Y>148 && Y<167)||(X-xSchuiving>20000 && X-xSchuiving<130000 && Y>148000 && Y<167000))) {// standplaats
						//results[17] = temp[19];
						if(temp[19].equals("Standplaats")) {
							
						}else {
							voegToe(X, Y, temp[19], "standplaats");
						}
					} else if (((X-xSchuiving ==119.0000 && temp[4].equals("189.5000") && (bordSpec ==1||bordSpec==0))||							
							(bordSpec != 1 &&(X-xSchuiving ==55.0000||X-xSchuiving ==35.0000||X-xSchuiving ==94.0000) && temp[4].equals("144.0000")))
							||((X-xSchuiving ==119000.0000 && (temp[4].equals("189500.0000")||temp[4].equals("189500.00")) && (bordSpec ==1||bordSpec==0))||
									(bordSpec != 1 && (X-xSchuiving ==55000.0000 && (temp[4].equals("144000.0000")||temp[4].equals("144000.00")))
									))
							) {// letterhoogte
						try {
						String[] tempSplit = new String[2];
						if(temp[19]!= null) {
							temp[19] =temp[19].replaceAll("/ ", "/");
						}
						String[] tempSplit3 = temp[19].split(" ");
						if(tempSplit3.length>0) {
						tempSplit[0] = tempSplit3[0];
						}else {
							tempSplit[0]=temp[19];
						}
						if(tempSplit3.length>1) {
							tempSplit = tempSplit3;
						}
						
						try {
						String[] tempSplit2 = new String[2];
						if(tempSplit != null && tempSplit.length>0) {
							if(tempSplit[0].contains("/")) {
							tempSplit2 = tempSplit[0].split("/");
							}else if(tempSplit[0].contains("-")) {
								tempSplit2 = tempSplit[0].split("-");
							}else {
								tempSplit2[0] = tempSplit[0];
							}
						}
						
						if (tempSplit2 != null &&tempSplit2.length>0 && tempSplit2[0] != null ) {
							if(isNumeric2(tempSplit2[0])) {
							results[18] = tempSplit2[0];
							}else {
								String part1="";
								String part2 = "";
								for(int x=0;x<tempSplit2[0].length();x++) {
									if(isNumeric2(tempSplit2[0].substring(x, x+1))){
										part1=part1+tempSplit2[0].substring(x, x+1);
									}else {
										part2 = tempSplit2[0].substring(x);
										break;
									}
								}
								results[18] = part1;
								
								tempSplit2[1] = part2+tempSplit2[1];
							}
						}
						
						if (tempSplit2.length>1 && tempSplit2[1] != null) {
							if(isNumeric2(tempSplit2[1])) {
								results[19] = tempSplit2[1];
							}else {
								String part1="";
								String part2 = "";
								for(int x=0;x<tempSplit2[1].length();x++) {
									if(x+1<tempSplit2[1].length() &&isNumeric2(tempSplit2[1].substring(x, x+1))){
										part1=part1+tempSplit2[1].substring(x, x+1);
									}else {
										part2 = tempSplit2[1].substring(x);
										break;
									}
								}
								results[19] = part1;
								
								if(tempSplit[1] == null) {
									tempSplit[1] = part2;
								}else {
								tempSplit[1] = part2+tempSplit[1];
								}
							}
							
							
							
						}
						}catch(Exception e) {
							System.out.println("test 5: "+e.getMessage()+"   "+pair.getKey());
						}
						
						try {
							try {
						if(tempSplit.length>1 && tempSplit[1]!= null) {
							if(tempSplit[1].contains("Rood")) {
								tempSplit[1] =tempSplit[1].replaceAll("Rood", "");
								results[22]="Rood";
							}
							if(tempSplit[1].contains("rood")) {
								tempSplit[1] = tempSplit[1].replaceAll("rood", "");
								results[22]="rood";
							}
							if(tempSplit[1].contains("ROOD")) {
								tempSplit[1] = tempSplit[1].replaceAll("ROOD", "");
								results[22]="ROOD";
							}
							if(tempSplit[1].contains("Goud")) {
								tempSplit[1] =tempSplit[1].replaceAll("Goud", "");
								if(results[22]==null) {
								results[22]="Goud";
								}else {
									results[22]=results[22]+" Goud";
								}
							}
							if(tempSplit[1].contains("goud")) {
								tempSplit[1] =tempSplit[1].replaceAll("goud", "");
								if(results[22]==null) {
								results[22]="goud";
								}else {
									results[22]=results[22]+" goud";
								}
							}
							if(tempSplit[1].contains("Zilver")) {
								tempSplit[1] =tempSplit[1].replaceAll("Zilver", "");
								if(results[22]==null) {
								results[22]="Zilver";
								}else {
									results[22]=results[22]+" Zilver";
								}
							}
							if(tempSplit[1].contains("zilver")) {
								tempSplit[1] =tempSplit[1].replaceAll("zilver", "");
								if(results[22]==null) {
								results[22]="zilver";
								}else {
									results[22]=results[22]+" zilver";
								}
							}
							if(tempSplit[1].contains("wit")) {
								tempSplit[1] =tempSplit[1].replaceAll("wit", "");
								if(results[22]==null) {
								results[22]="wit";
								}else {
									results[22]=results[22]+" wit";
								}
							}
							if(tempSplit[1].contains("zwart")) {
								tempSplit[1] =tempSplit[1].replaceAll("zwart", "");
								if(results[22]==null) {
								results[22]="zwart";
								}else {
									results[22]=results[22]+" zwart";
								}
							}
							
						}
							}catch(Exception e) {
								System.out.println("test2: "+e.getMessage());
							}
						if (tempSplit.length > 2) {
							
							if (tempSplit[2] != null) {
								for(int x=2;x<tempSplit.length;x++) {
									if(results[22]==null) {
										results[22] = tempSplit[x];
									}else {
										results[22] = results[22] +" "+tempSplit[x];
									}
								}
								if(results[22].contains("Dd")) {
									results[22]=results[22].replaceAll("Dd", "");
									tempSplit[1]=tempSplit[1]+"/Dd";
									
								}
								if(results[22].contains("Dn")) {
									results[22]=results[22].replaceAll("Dn", "");
									tempSplit[1]=tempSplit[1]+"/Dn";
									
								}
								if(results[22].contains("Cz")) {
									results[22]=results[22].replaceAll("Cz", "");
									tempSplit[1]=tempSplit[1]+"/Cz";
									
								}
								if(results[22].contains("Uu")) {
									results[22]=results[22].replaceAll("Uu", "");
									tempSplit[1]=tempSplit[1]+"/Uu";
									
								}
								results[22]=results[22].replaceAll("mm", "").strip();
								results[22]=results[22].replaceAll("MM.", "").strip();
								results[22]=results[22].replaceAll("-", "").strip();
								results[22]=results[22].replaceAll("[()]", "").strip();
								if(isNumeric2(results[22].replaceAll("/", ""))) {
									results[22]="";
								}
								//results[22]=results[22].strip();
								//System.out.println(results[22]);
								
								//results[22] = tempSplit[tempSplit.length - 1];
							}
						}
						if(results[22]!= null && (results[22].equals("()")||results[22].startsWith("/"))) {
							results[22]="";
						}
						}catch(Exception e) {
							System.out.println("test: "+e.getMessage());
						}
						try {
						if (tempSplit.length > 1 && tempSplit[1] != null) {
							
							if (tempSplit[1] != null) {
								if(tempSplit[1].contains("/")) { 
									String[] tmp =tempSplit[1].split("/");
									if(tmp.length>0) {
									results[20] = tmp[0];
									if(tmp.length>1) {
									results[21] = tmp[1];
									}
									}
								}else if(tempSplit[1].contains("-")) { 
									String[] tmp =tempSplit[1].split("-");
									if(tmp.length>0) {
									results[20] = tmp[0];
									if(tmp.length>1) {
									results[21] = tmp[1];
									}
									}
								}else {
									results[20] = tempSplit[1];
								}
								results[20] = results[20].replaceAll("[(mm)]", "").replaceAll("mm", "");
							}
						}
						if(results[20]!= null && results[20].contains("null")) {
							results[20]=results[20].substring(0, results[20].indexOf("null"));
						}
						if(results[20]!= null && results[20].contains(";10")) {
							results[20]=results[20].replaceAll(";10", "");
						}
						if(results[20]!= null && results[20].contains("()")) {
							
							results[20]=results[20].replaceAll("[()]", "");
							
						}
						if(results[21]!= null && isNumeric2(results[21].replaceAll("/", ""))) {
							results[21]="";
						}
						}catch(Exception e) {
							System.out.println("test3: "+e.getMessage());
						}
						}catch(Exception e) {
							System.out.println("test 8: "+e.getMessage());
						}
					} else {
						if(bordSpec ==1||bordSpec==0) {
							if(temp[3] != null && temp[4] != null && !temp[3].isEmpty()&& !temp[4].isEmpty()) {
						X = Double.valueOf(temp[3]);
						Y = Double.valueOf(temp[4]);
						if ((X-xSchuiving > 21*1000 && X-xSchuiving < 27*1000)||(X-xSchuiving > 21 && X-xSchuiving < 27)) {// linker rij bouwjaar
							if ((Y > 163 && Y <= 188)||(Y > 163000 && Y <= 188000)) { // 1
								voegToe(X, Y, temp[19], "bj_1");
							} else if ((Y > 137 && Y <= 162)||(Y > 137000 && Y <= 162000)) { // 3
								voegToe(X, Y, temp[19], "bj_3");
							} else if ((Y > 111 && Y <= 135)||(Y > 111000 && Y <= 135000)) { // 5
								voegToe(X, Y, temp[19], "bj_5");
							} else if ((Y > 84 && Y <= 109)||(Y > 84000 && Y <= 109000)) { // 7
								voegToe(X, Y, temp[19], "bj_7");
							} else if ((Y > 58 && Y <= 83)||(Y > 58000 && Y <= 83000)) { // 9
								voegToe(X, Y, temp[19], "bj_9");
							} else if ((Y > 32 && Y <= 58)||(Y > 32000 && Y <= 58000)) { // 11
								voegToe(X, Y, temp[19], "bj_11");
							}

						} else if ((X-xSchuiving > 190 && X-xSchuiving < 198)||(X-xSchuiving > 190000 && X-xSchuiving < 198000)) {// rechter rij bouwjaar
							if ((Y > 163 && Y < 188)||(Y > 163000 && Y < 188000)) { // 2
								voegToe(X, Y, temp[19], "bj_2");
							} else if ((Y > 137 && Y <= 162)||(Y > 137000 && Y <= 162000)) { // 4
								voegToe(X, Y, temp[19], "bj_4");
							} else if ((Y > 111 && Y <= 135)||(Y > 111000 && Y <= 135000)) { // 6
								voegToe(X, Y, temp[19], "bj_6");
							} else if ((Y > 84 && Y <= 109)||(Y > 84000 && Y <= 109000)) { // 8
								voegToe(X, Y, temp[19], "bj_8");
							} else if ((Y > 58 && Y <= 83)||(Y > 58000 && Y <= 83000)) { // 10
								voegToe(X, Y, temp[19], "bj_10");
							} else if ((Y > 32 && Y <= 58)||(Y > 32000 && Y <= 58000)) { // 12
								voegToe(X, Y, temp[19], "bj_12");
							}
						} else if ((X-xSchuiving > 28 && X-xSchuiving < 38)||(X-xSchuiving > 28000 && X-xSchuiving < 38000)) {// linker rij kleur
							if(temp[19].equals(".")== false) {
								if ((Y > 163 && Y <= 175)||(Y > 163000 && Y <= 175000)) { // 1
									voegToe(X, Y, temp[19], "kleur_1_b");
								} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 1
									voegToe(X, Y, temp[19], "kleur_1_a");
								} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 3
									voegToe(X, Y, temp[19], "kleur_3_b");
								} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 3
									voegToe(X, Y, temp[19], "kleur_3_a");
								} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 5
									voegToe(X, Y, temp[19], "kleur_5_b");
								} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 5
									voegToe(X, Y, temp[19], "kleur_5_a");
								} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 7
									voegToe(X, Y, temp[19], "kleur_7_b");
								} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 7
									voegToe(X, Y, temp[19], "kleur_7_a");
								} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 9
									voegToe(X, Y, temp[19], "kleur_9_b");
								} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 9
									voegToe(X, Y, temp[19], "kleur_9_a");
								} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 11
									voegToe(X, Y, temp[19], "kleur_11_b");
								} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 11
									voegToe(X, Y, temp[19], "kleur_11_a");
								}
							}

						} else if ((X-xSchuiving > 190 && X-xSchuiving < 198)||(X-xSchuiving > 190000 && X-xSchuiving < 198000)) {// rechter rij kleur
							if(temp[19].equals(".")== false) {
								if ((Y > 163 && Y < 175)||(Y > 163000 && Y < 175000)) { // 2
									voegToe(X, Y, temp[19], "kleur_2_a");
								} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 2
									voegToe(X, Y, temp[19], "kleur_2_b");
								} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 4
									voegToe(X, Y, temp[19], "kleur_4_a");
								} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 4
									voegToe(X, Y, temp[19], "kleur_4_b");
								} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 6
									voegToe(X, Y, temp[19], "kleur_6_a");
								} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 6
									voegToe(X, Y, temp[19], "kleur_6_b");
								} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 8
									voegToe(X, Y, temp[19], "kleur_8_a");
								} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 8
									voegToe(X, Y, temp[19], "kleur_8_b");
								} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 10
									voegToe(X, Y, temp[19], "kleur_10_a");
								} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 10
									voegToe(X, Y, temp[19], "kleur_10_b");
								} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 12
									voegToe(X, Y, temp[19], "kleur_12_a");
								} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 12
									voegToe(X, Y, temp[19], "kleur_12_b");
								}
							}
						} else if ((X-xSchuiving > 38 && X-xSchuiving < 49)||(X-xSchuiving > 38000 && X-xSchuiving < 49000)) {// linker rij afstand
							if ((Y > 163 && Y <= 188)||(Y > 163000 && Y <= 188000)) { // 1
								voegToe(X, Y, temp[19], "afstand_1");
							} else if ((Y > 137 && Y <= 162)||(Y > 137000 && Y <= 162000)) { // 3
								voegToe(X, Y, temp[19], "afstand_3");
							} else if ((Y > 111 && Y <= 135)||(Y > 111000 && Y <= 135000)) { // 5
								voegToe(X, Y, temp[19], "afstand_5");
							} else if ((Y > 84 && Y <= 109)|| (Y > 84000 && Y <= 109000)) { // 7
								voegToe(X, Y, temp[19], "afstand_7");
							} else if ((Y > 58 && Y <= 83)||(Y > 58000 && Y <= 83000)) { // 9
								voegToe(X, Y, temp[19], "afstand_9");
							} else if ((Y > 32 && Y <= 58)||(Y > 32000 && Y <= 58000)) { // 11
								voegToe(X, Y, temp[19], "afstand_11");
							}

						} else if ((X-xSchuiving > 190 && X-xSchuiving < 198)||(X-xSchuiving > 190000 && X-xSchuiving < 198000)) {// rechter rij afstand
							if ((Y > 163 && Y < 188)||(Y > 163000 && Y < 188000)) { // 2
								voegToe(X, Y, temp[19], "afstand_2");
							} else if ((Y > 137 && Y <= 162)||(Y > 137000 && Y <= 162000)) { // 4
								voegToe(X, Y, temp[19], "afstand_4");
							} else if ((Y > 111 && Y <= 135)||(Y > 111000 && Y <= 135000)) { // 6
								voegToe(X, Y, temp[19], "afstand_6");
							} else if ((Y > 84 && Y <= 109)||(Y > 84000 && Y <= 109000)) { // 8
								voegToe(X, Y, temp[19], "afstand_8");
							} else if ((Y > 58 && Y <= 83)||(Y > 58000 && Y <= 83000)) { // 10
								voegToe(X, Y, temp[19], "afstand_10");
							} else if ((Y > 32 && Y <= 58)||(Y > 32000 && Y <= 58000)) { // 12
								voegToe(X, Y, temp[19], "afstand_12");
							}
						} else if (((X-xSchuiving > 63 && X-xSchuiving < 108)||(X-xSchuiving > 63000 && X-xSchuiving < 108000))&& !temp[27].contentEquals("CAMTAG")&& temp[24].isBlank()) {// linker rij bord
							
							if ((Y > 163 && Y <= 175)||(Y > 163000 && Y <= 175000)) { // 1
								voegToe(X, Y, temp[19], "bord_1_b");
							} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 1
								voegToe(X, Y, temp[19], "bord_1_a");
							} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 3
								voegToe(X, Y, temp[19], "bord_3_b");
							} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 3
								voegToe(X, Y, temp[19], "bord_3_a");
							} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 5
								voegToe(X, Y, temp[19], "bord_5_b");
							} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 5
								voegToe(X, Y, temp[19], "bord_5_a");
							} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 7
								voegToe(X, Y, temp[19], "bord_7_b");
							} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 7
								voegToe(X, Y, temp[19], "bord_7_a");
							} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 9
								voegToe(X, Y, temp[19], "bord_9_b");
							} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 9
								voegToe(X, Y, temp[19], "bord_9_a");
							} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 11
								voegToe(X, Y, temp[19], "bord_11_b");
							} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 11
								voegToe(X, Y, temp[19], "bord_11_a");
							}
							
						} else if (((X-xSchuiving > 110 && X-xSchuiving < 155)||(X-xSchuiving > 110000 && X-xSchuiving < 155000)) && !temp[27].contentEquals("CAMTAG") && temp[24].isBlank()) {// rechter rij bord
							
							if ((Y > 163 && Y < 175)||(Y > 163000 && Y < 175000)) { // 2
								voegToe(X, Y, temp[19], "bord_2_a");
							} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 2
								voegToe(X, Y, temp[19], "bord_2_b");
							} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 4
								voegToe(X, Y, temp[19], "bord_4_a");
							} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 4
								voegToe(X, Y, temp[19], "bord_4_b");
							} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 6
								voegToe(X, Y, temp[19], "bord_6_a");
							} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 6
								voegToe(X, Y, temp[19], "bord_6_b");
							} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 8
								voegToe(X, Y, temp[19], "bord_8_a");
							} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 8
								voegToe(X, Y, temp[19], "bord_8_b");
							} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 10
								voegToe(X, Y, temp[19], "bord_10_a");
							} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 10
								voegToe(X, Y, temp[19], "bord_10_b");
							} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 12
								voegToe(X, Y, temp[19], "bord_12_a");
							} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 12
								voegToe(X, Y, temp[19], "bord_12_b");
							}
							
						} else if ((X-xSchuiving > 50 && X-xSchuiving < 109)||(X-xSchuiving > 50000 && X-xSchuiving < 109000)) {// linker rij TAG
							
							if(temp[27].equals("CAMTAG") && !temp[19].isBlank()) {
								temp[24] = temp[19];								
							}
							if ((Y > 162 && Y <= 188)||(Y > 162000 && Y <= 188000)) { // 1
								voegToe(X, Y, temp[24], "TAG_1");
							} else if ((Y > 134 && Y <= 162)||(Y > 134000 && Y <= 162000)) { // 3
								voegToe(X, Y, temp[24], "TAG_3");
							} else if ((Y > 109 && Y <= 134)||(Y > 109000 && Y <= 134000)) { // 5
								voegToe(X, Y, temp[24], "TAG_5");
							} else if ((Y > 84 && Y <= 109)||(Y > 84000 && Y <= 109000)) { // 7
								voegToe(X, Y, temp[24], "TAG_7");
							} else if ((Y > 57 && Y <= 84)||(Y > 57000 && Y <= 84000)) { // 9
								voegToe(X, Y, temp[24], "TAG_9");
							} else if ((Y > 32 && Y <= 57)||(Y > 32000 && Y <= 57000)) { // 11
								voegToe(X, Y, temp[24], "TAG_11");
							}
						} else if ((X-xSchuiving >= 109 && X-xSchuiving < 169)||(X-xSchuiving >= 109000 && X-xSchuiving < 169000)) {// rechter rij TAG
							
							if(temp[27].contentEquals("CAMTAG")&& !temp[19].isBlank()) {
								temp[24] = temp[19];							
							}
							if ((Y > 162 && Y < 188)||(Y > 162000 && Y < 188000)) { // 2
								voegToe(X, Y, temp[24], "TAG_2");
							} else if ((Y > 134 && Y <= 162)||(Y > 134000 && Y <= 162000)) { // 4
								voegToe(X, Y, temp[24], "TAG_4");
							} else if ((Y > 109 && Y <= 134)||(Y > 109000 && Y <= 134000)) { // 6
								voegToe(X, Y, temp[24], "TAG_6");
							} else if ((Y > 84 && Y <= 109)||(Y > 84000 && Y <= 109000)) { // 8
								voegToe(X, Y, temp[24], "TAG_8");
							} else if ((Y > 57 && Y <= 84)||(Y > 57000 && Y <= 84000)) { // 10
								voegToe(X, Y, temp[24], "TAG_10");
							} else if ((Y > 32 && Y <= 57)||(Y > 32000 && Y <= 57000)) { // 12
								voegToe(X, Y, temp[24], "TAG_12");
							}
						}
						
						if ((X-xSchuiving > 20000 && X-xSchuiving < 47000 && Y > 0 && Y < 28000)||
								(X-xSchuiving > 20 && X-xSchuiving < 47 && Y > 0 && Y < 28)) {// Art.code plaatsing
							
							if(verschovenBord == true || temp[20].equals("2")) {
								voegToe(X, Y, temp[19].strip(), "aant.");
							}else {
							voegToe(X, Y, temp[19].strip(), "Art.code plaatsing");
							}
						}
						if ((X-xSchuiving > 47000 && X-xSchuiving < 57000 && Y > 0 && Y < 28000)||
								(X-xSchuiving > 47 && X-xSchuiving < 57 && Y > 0 && Y < 28)) {// aant.
							if(verschovenBord == true || temp[20].equals("2")) {
								voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
							}else {
							voegToe(X, Y, temp[19].strip(), "aant.");
							}
						}
						if ((X-xSchuiving > 57000 && X-xSchuiving < 168000 && Y > 0 && Y < 28000)||
								(X-xSchuiving > 57 && X-xSchuiving < 168 && Y > 0 && Y < 28)) {// omschrijving werkzaamheden
							voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
						}
						if ((X-xSchuiving > 62000 && X-xSchuiving < 76000 && Y > 208000 && Y < 260000 )||
								(X-xSchuiving > 62 && X-xSchuiving < 76 && Y > 208 && Y < 260 )) {// Artikel
							if(verschovenBord == true) {
								voegToe(X, Y, temp[19].strip(), "Nw");
							}else {
							voegToe(X, Y, temp[19].strip(), "artikel");
							}
						}
						if ((X-xSchuiving > 80000 && X-xSchuiving < 85000 && Y > 208000 && Y < 260000)||
								(X-xSchuiving > 80 && X-xSchuiving < 85 && Y > 208 && Y < 260)) {// Nw
							if(verschovenBord == true) {
								voegToe(X, Y, temp[19].strip(), "Tot");
							}else {
							voegToe(X, Y, temp[19].strip(), "Nw");
							}
						}
						if ((X-xSchuiving > 85000 && X-xSchuiving < 90000 && Y > 208000 && Y < 260000)||
								(X-xSchuiving > 85 && X-xSchuiving < 90 && Y > 208 && Y < 260)) {// Tot
							voegToe(X, Y, temp[19].strip(), "Tot");
						}
						if ((X-xSchuiving > 90000 && X-xSchuiving < 164000 && Y > 208000 && Y < 260000)||
								(X-xSchuiving > 90 && X-xSchuiving < 164 && Y > 208 && Y < 260)) {// Omschrijving
							voegToe(X, Y, temp[19].strip(), "Omschrijving");
						}
						if ((X-xSchuiving > 164000 && X-xSchuiving < 176000 && Y > 208000 && Y < 260000)||
								(X-xSchuiving > 164 && X-xSchuiving < 176 && Y > 208 && Y < 260)) {// Bouwjr
							voegToe(X, Y, temp[19].strip(), "Bouwjr");
						}
							}
					}else {
						//BORDSPEC niet 1
						if(temp[3] != null && temp[4] != null && !temp[3].isEmpty()&& !temp[4].isEmpty()) {
							X = Double.valueOf(temp[3]);
							Y = Double.valueOf(temp[4]);
							if(bordSpec==2) {
								if ((X-xSchuiving > 21 && X-xSchuiving < 40 && Y > 0 && Y < 32)||(X-xSchuiving > 21000 && X-xSchuiving < 35000 && Y > 0 && Y < 32000)) {// Art.code plaatsing
									// System.out.println(temp[19]);
									
									voegToe(X, Y, temp[19].strip(), "Art.code plaatsing");
									
								}
								if((X-xSchuiving > 40 && X-xSchuiving < 48 && Y > 0 && Y < 32)||(X-xSchuiving > 40000 && X-xSchuiving < 48000 && Y > 0 && Y < 32000)) {// Art.code plaatsing
									if(verschovenBord == true) {
										voegToe(X, Y, temp[19].strip(), "aant.");
									}
								}
								if ((X-xSchuiving > 47 && X-xSchuiving < 57 && Y > 0 && Y < 28)||(X-xSchuiving > 47000 && X-xSchuiving < 57000 && Y > 0 && Y < 28000)) {// aant.
									if(verschovenBord == true) {
										voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
									}else {
									voegToe(X, Y, temp[19].strip(), "aant.");
									}
								}
								if ((X-xSchuiving > 57 && X-xSchuiving < 168 && Y > 0 && Y < 28)||(X-xSchuiving > 57000 && X-xSchuiving < 168000 && Y > 0 && Y < 28000)) {// omschrijving werkzaamheden
									voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
								}
							}else {
								if ((X-xSchuiving > 21 && X-xSchuiving < 47 && Y > 0 && Y < 28)||(X-xSchuiving > 21000 && X-xSchuiving < 47000 && Y > 0 && Y < 28000)) {// Art.code plaatsing
									// System.out.println(temp[19]);
									if(verschovenBord == true) {
										voegToe(X, Y, temp[19].strip(), "aant.");
									}else {
										voegToe(X, Y, temp[19].strip(), "Art.code plaatsing");
									}
								}
								if ((X-xSchuiving > 47 && X-xSchuiving < 57 && Y > 0 && Y < 28)||(X-xSchuiving > 47000 && X-xSchuiving < 57000 && Y > 0 && Y < 28000)) {// aant.
									if(verschovenBord == true) {
										voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
									}else {
									voegToe(X, Y, temp[19].strip(), "aant.");
									}
								}
								if ((X-xSchuiving > 57 && X-xSchuiving < 168 && Y > 0 && Y < 28)||(X-xSchuiving > 57000 && X-xSchuiving < 168000 && Y > 0 && Y < 28000)) {// omschrijving werkzaamheden
									voegToe(X, Y, temp[19].strip(), "omschrijving werkzaamheden");
								}
							}
							if ((X-xSchuiving > 20 && X-xSchuiving < 40 && Y > 169 && Y < 236)||(X-xSchuiving > 20000 && X-xSchuiving < 40000 && Y > 169000 && Y < 236000)) {// Artikel
								
									voegToe(X, Y, temp[19].strip(), "artikel");
								
								
							}
							if ((X-xSchuiving > 40 && X-xSchuiving < 45 && Y > 169 && Y < 236)||(X-xSchuiving > 40000 && X-xSchuiving < 45000 && Y > 169000 && Y < 236000)) {// Nw
								
									voegToe(X, Y, temp[19].strip(), "Nw");
								
								
							}
							if ((X-xSchuiving > 45 && X-xSchuiving < 58.5 && Y > 169 && Y < 236)||(X-xSchuiving > 45000 && X-xSchuiving < 58500 && Y > 169000 && Y < 236000)) {// Tot
								
								voegToe(X, Y, temp[19].strip(), "Tot");
								
							}
							if ((X-xSchuiving > 58.5 && X-xSchuiving < 182 && Y > 169 && Y < 236)||(X-xSchuiving > 58500 && X-xSchuiving < 182000 && Y > 169000 && Y < 236000)) {// Omschrijving
								voegToe(X, Y, temp[19].strip(), "Omschrijving");
							}
							if ((X-xSchuiving > 182 && X-xSchuiving < 197 && Y > 169 && Y < 236)||(X-xSchuiving > 182000 && X-xSchuiving < 197000 && Y > 169000 && Y < 236000)) {// Bouwjr
								voegToe(X, Y, temp[19].strip(), "Bouwjr");
							}
							if(bordSpec == -1||bordSpec == 2) {
								if ((X-xSchuiving > 70 && X-xSchuiving < 170 && Y > 242 && Y < 266)||(X-xSchuiving > 70000 && X-xSchuiving < 170000 && Y > 242000 && Y < 266000)) {// Opmerkingen
									voegToe(X, Y, temp[19].strip(), "Opmerkingen");
								}
								if ((X-xSchuiving > 185 && X-xSchuiving < 186 && Y > 27.5&& Y < 28.5)||(X-xSchuiving > 185000 && X-xSchuiving < 186000 && Y > 27500 && Y < 28500)) {								
									if(results[2]==null || results[2].isEmpty()) {
										results[2] = temp[19]; // getekend
									}
								}
								if ((X-xSchuiving > 182 && X-xSchuiving < 183 && Y > 21.9 && Y < 22.1)||(X-xSchuiving > 182000 && X-xSchuiving < 183000 && Y > 21900 && Y < 22100)) {								
									if(results[3]==null || results[3].isEmpty()) {
										results[3] = temp[19]; // datum
									}
								}
								
								if ((X-xSchuiving > 181 && X-xSchuiving < 184 && Y > 14 && Y < 18)||(X-xSchuiving > 181000 && X-xSchuiving < 184000 && Y > 14000 && Y < 18000)) {								
									if(results[4]==null || results[4].isEmpty()) {
										results[4] = temp[19]; // gezien
									}
								}
								

							}else if(bordSpec == 3) {
								if ((X-xSchuiving > 70 && X-xSchuiving < 170 && Y > 242 && Y < 261)||(X-xSchuiving > 70000 && X-xSchuiving < 170000 && Y > 242000 && Y < 261000)) {// Opmerkingen
									voegToe(X, Y, temp[19].strip(), "Opmerkingen");
								}else if ((X-xSchuiving > 70 && X-xSchuiving < 170 && Y > 261 && Y < 266)||(X-xSchuiving > 70000 && X-xSchuiving < 170000 && Y > 261000 && Y < 266000)) {// Consulkring
									voegToe(X, Y, temp[19].strip(), "Consulkring");
								}
								
							}else if(bordSpec == -2) {
								if ((X-xSchuiving > 70 && X-xSchuiving < 170 && Y > 242 && Y < 266)||(X-xSchuiving > 70000 && X-xSchuiving < 170000 && Y > 242000 && Y < 266000)) {// Opmerkingen
									voegToe(X, Y, temp[19].strip(), "Opmerkingen");
								}
								
								if ((X-xSchuiving > 187 && X-xSchuiving < 198 && Y > 32 && Y < 142)||(X-xSchuiving > 187000 && X-xSchuiving < 198000 && Y > 32000 && Y < 142000)) {// bj
									voegToe(X, Y, temp[19].strip(), "bj_1");
								}
							}
							
							
						}
					}
					}
					}catch(Exception e) {
						System.out.println("test test: "+e.getMessage());
					}
				} else if (temp[2].equals("VOETTUN")==false &&!temp[2].equals("INFOP")&&((temp[2].endsWith("KP") || temp[2].endsWith("OP") || temp[2].endsWith("OK")
						||temp[2].endsWith("UML")||temp[2].endsWith("CC")|| temp[2].endsWith("CC CAP")||temp[2].endsWith("UZSTR")||temp[2].endsWith("LHK")||temp[2].endsWith("RHK")
						||temp[2].startsWith("ACAD")|| temp[2].startsWith("C07")||temp[2].startsWith("SVRAUTO")||temp[2].endsWith("KK") || temp[2].endsWith("UN") || temp[2].endsWith("PO")))) {
					if(temp[3] != null && temp[4] != null && !temp[3].isEmpty()&& !temp[4].isEmpty()) {
						try {
					double X = Double.valueOf(temp[3]);
					double Y = Double.valueOf(temp[4]);
					
					
					// temp[2].endsWith("OP") ||temp[2].endsWith("UN") is speciaal teken
					// temp[2].endsWith("KK") is cijfer
					// temp[2].endsWith("KP") is hoofdletter
					// temp[2].endsWith("OK") is kleine letter
					String teken = "";
					if (temp[2].startsWith("ACAD")) {
						teken = "";
					}else if (temp[2].startsWith("C07")||temp[2].startsWith("SVRAUTO")) {
						teken = "";
					}else if (temp[2].endsWith("UML")) {
						teken = "\"";
					}
					else if (temp[2].endsWith("UZSTR")) {
						teken = "/";
					}else if (temp[2].endsWith("LHK")) {
						teken = "(";
					}else if (temp[2].endsWith("RHK")) {
						teken = ")";
					}
					else if (temp[2].endsWith("CC")|| temp[2].endsWith("CC CAP")) {
						teken = "";
					}
					else if (temp[2].endsWith("KP") || temp[2].endsWith("KK")) {
						teken = temp[2].substring(2, temp[2].length() - 2);
					} else if (temp[2].endsWith("OK")) {
						teken = temp[2].substring(2, temp[2].length() - 2).toLowerCase();
					} else if (temp[2].endsWith("OP") || (temp[2].endsWith("UN"))) {
						teken = temp[2].substring(2, temp[2].length() - 2);
						if (teken.equals("K")) {
							teken = "-";
						} else if (teken.equals("P")) {
							teken = ".";
						}else if (teken.equals("NK")||teken.equals("ZK")||teken.equals("NP")||teken.equals("SCZPUN")||teken.equals("SEZPUN")||teken.equals("SUZPUN")) {
							teken = "";
						} else {
							teken = "";
							//System.out.println(pair.getKey()+"     Speciaal teken: " + temp[2] + "  wordt niet behandeld");
						}
					} else if (temp[2].endsWith("PO") && temp[2].length()>3) {
						teken = temp[2].substring(2, temp[2].length() - 2);
						if (teken.equals("A")) {
							teken = "'";
						}else if (teken.equals("O")) {
							teken = "";
						}else if (teken.equals("NA")) {
							teken = "";
						}else if (teken.equals("ZA")) {
							teken = "";
						} else {
							teken = "";
							//System.out.println(pair.getKey()+"       Speciaal teken: " + temp[2] + "  wordt niet behandeld");
						}
					}
					if(bordSpec == 1||bordSpec==0) {
					if ((X > 63 && X < 108)||(X > 63000 && X < 108000)) {// linker rij bord
						if ((Y > 163 && Y <= 175)||(Y > 163000 && Y <= 175000)) { // 1
							voegToe(X, Y, teken, "bord_1_b");
						} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 1
							voegToe(X, Y, teken, "bord_1_a");
						} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 3
							voegToe(X, Y, teken, "bord_3_b");
						} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 3
							voegToe(X, Y, teken, "bord_3_a");
						} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 5
							voegToe(X, Y, teken, "bord_5_b");
						} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 5
							voegToe(X, Y, teken, "bord_5_a");
						} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 7
							voegToe(X, Y, teken, "bord_7_b");
						} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 7
							voegToe(X, Y, teken, "bord_7_a");
						} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 9
							voegToe(X, Y, teken, "bord_9_b");
						} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 9
							voegToe(X, Y, teken, "bord_9_a");
						} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 11
							voegToe(X, Y, teken, "bord_11_b");
						} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 11
							voegToe(X, Y, teken, "bord_11_a");
						}
					} else if ((X > 110 && X < 155)||(X > 110000 && X < 155000)) {// rechter rij bord
						if ((Y > 163 && Y < 175)||(Y > 163000 && Y < 175000)) { // 2
							voegToe(X, Y, teken, "bord_2_a");
						} else if ((Y > 175 && Y < 188)||(Y > 175000 && Y < 188000)) { // 2
							voegToe(X, Y, teken, "bord_2_b");
						} else if ((Y > 137 && Y <= 149.5)||(Y > 137000 && Y <= 149500)) { // 4
							voegToe(X, Y, teken, "bord_4_a");
						} else if ((Y > 149.5 && Y < 162)||(Y > 149500 && Y < 162000)) { // 4
							voegToe(X, Y, teken, "bord_4_b");
						} else if ((Y > 111 && Y <= 122)||(Y > 111000 && Y <= 122000)) { // 6
							voegToe(X, Y, teken, "bord_6_a");
						} else if ((Y > 122 && Y < 135)||(Y > 122000 && Y < 135000)) { // 6
							voegToe(X, Y, teken, "bord_6_b");
						} else if ((Y > 84 && Y <= 97)||(Y > 84000 && Y <= 97000)) { // 8
							voegToe(X, Y, teken, "bord_8_a");
						} else if ((Y > 97 && Y < 109)||(Y > 97000 && Y < 109000)) { // 8
							voegToe(X, Y, teken, "bord_8_b");
						} else if ((Y > 58 && Y <= 71)||(Y > 58000 && Y <= 71000)) { // 10
							voegToe(X, Y, teken, "bord_10_a");
						} else if ((Y > 71 && Y < 83)||(Y > 71000 && Y < 83000)) { // 10
							voegToe(X, Y, teken, "bord_10_b");
						} else if ((Y > 32 && Y <= 44)||(Y > 32000 && Y <= 44000)) { // 12
							voegToe(X, Y, teken, "bord_12_a");
						} else if ((Y > 44 && Y < 58)||(Y > 44000 && Y < 58000)) { // 12
							voegToe(X, Y, teken, "bord_12_b");
						}
					}
					}else if (bordSpec == 2) {
						if ((X+xSchuiving > 40 && X+xSchuiving < 180 )||(X+xSchuiving > 40000 && X+xSchuiving < 180000 )) {
							if ((Y > 50 && Y < 120)||(Y > 50000 && Y < 120000)) {
								
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_a");
								
							}
						}
					}else if(bordSpec == 3) {
						//paddestoel
						
						
						if ((X+xSchuiving > 55 && X+xSchuiving < 100 )||(X+xSchuiving > 55000 && X+xSchuiving < 100000 )) {
							if ((Y > 86 && Y < 107)||(Y > 86000 && Y < 107000)) {
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_a");
							}else if ((Y > 42 && Y < 64)||(Y > 42000 && Y < 64000)) {
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_b");
							}
						}
											
						
					}else if(bordSpec == -1) {
						//FRSPEC
						if ((X+xSchuiving > 20 && X+xSchuiving < 172 )||(X+xSchuiving > 20000 && X+xSchuiving < 172000 )) {
							if ((Y > 32 && Y < 142)||(Y > 32000 && Y < 142000)) {
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_a");
							}
						}
					}else if(bordSpec == -2) {
						//FLSPECI
						if ((X+xSchuiving > 61 && X+xSchuiving < 175 )||(X+xSchuiving > 61000 && X+xSchuiving < 175000 )) {
							if ((Y > 32 && Y < 142)||(Y > 32000 && Y < 142000)) {
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_a");
							}
						}
					}else if(bordSpec == 5) {
						
						if (X+xSchuiving > 0) {
							if (Y > 0 ) {
								if(teken.equals("*")||teken.equals("**")||teken.equals("***")||teken.equals("****")
										||teken.equals("*****")||teken.equals("******")||teken.equals("*******")) {
									voegToe(X, Y, teken, "kleur_1");
								}
								voegToe(X, Y, teken, "bord_1_a");
							}
						}
					}
					
					}catch(Exception e) {
						System.out.println("test test2: "+e.getMessage());
					}
					}
				} else if ((temp[3] != null && temp[4] != null && !temp[3].isEmpty() && !temp[4].isEmpty())||(temp[2].equals("Polyline"))) { // name is
																												// niet
																												// een
																												// van
																												// de
																												// standaard
					
					if(temp[2].equals("Polyline") )	{	
						//temp[20] == layer
						//temp[25] == closed -1 is open, 0 is closed
						if(firstPoly == true&&(temp[20].equals("87")||temp[20].equals("to-xx-bibcontour")) &&temp[25].equals("-1")) {
						
						firstPoly = false;
						voegToe(0, 0, "pijl Onbekend", "pijl_bord_onbekend");
						
						}
					}else {
					try {
					double X = Double.valueOf(temp[3]);
					double Y = Double.valueOf(temp[4]);
					String teken = temp[2];
					if ((teken.equals("PYLA") ||teken.equals("PYLB")||teken.equals("PYLU") || teken.equals("nbpijl")|| teken.equals("nbPIJL")
							|| teken.startsWith("BORD")|| teken.startsWith("BCBORD")|| teken.startsWith("TBORD")|| teken.startsWith("FBORD"))&&(!teken.startsWith("BORDBAL"))) {
						if(teken.startsWith("BCBORDFU")) {
							String tmp =teken.replaceFirst("BCBORDFU", "");
							if(tmp.startsWith("0")){
								teken = "pijl 90 graden (rechts)";
							}else if(tmp.startsWith("1")) {
								teken = "pijl 270 graden (links)";
							}
						}else if(teken.startsWith("BCBORDVN")) {
							String tmp =teken.replaceFirst("BCBORDVN", "");
							if(tmp.startsWith("0")){
								teken = "pijl 90 graden (rechts)";
							}else if(tmp.startsWith("1")) {
								teken = "pijl 270 graden (links)";
							}
						}else if(teken.startsWith("BCBORDVU")) {
							String tmp =teken.replaceFirst("BCBORDVU", "");
							if(tmp.startsWith("0")){
								teken = "pijl 90 graden (rechts)";
							}else if(tmp.startsWith("1")) {
								teken = "pijl 270 graden (links)";
							}
						}else if(teken.startsWith("BORDRING") ) {
							voegToe(X, Y, "RING", "bord_1_a");
						}else if(teken.startsWith("BORDUIT") ) {
							voegToe(X, Y, "UIT", "bord_1_a");
							teken = "pijl 45 graden (midden)";
						}
						else if(teken.startsWith("BCBORDKH") ||teken.startsWith("BCBORDVH") || teken.startsWith("BORDRUI")||teken.startsWith("BORDPUN1")||teken.startsWith("BORDSNR")||teken.startsWith("BORDKH")|| teken.startsWith("FBORD")|| teken.startsWith("BORDKN02")) {
							teken="";
						}else if(teken.startsWith("BORDWEG") ||teken.startsWith("BORDKN") ||teken.startsWith("BORDLN") ||teken.startsWith("BORDDATA") ||teken.startsWith("BORDPUN") ||teken.startsWith("BORDVU") ||teken.startsWith("BORDKNBL") ||teken.startsWith("BORDKNBR") ||teken.startsWith("BORDGH") || teken.startsWith("TBORD")||teken.startsWith("BCBORDGH") ||teken.startsWith("BORDID") ||teken.startsWith("BORDVN") ||teken.startsWith("BCBORDFN") ||teken.startsWith("BORDFN") ||teken.startsWith("BORDFUB") ||
								teken.startsWith("BCBORDFH") ||teken.startsWith("BORDFH") ||teken.startsWith("BORDLUB") || teken.startsWith("BORDGNB")|| teken.startsWith("BORDLNB") ||
								teken.startsWith("BCBORDLW")||teken.startsWith("BORDLW")||teken.startsWith("BORDLWB")||teken.startsWith("BORDKUB")||teken.startsWith("BORDFU0") ||teken.startsWith("BORDFU1") || teken.startsWith("BORDGUB")) {
							teken= ""; //achterkant bord
						}else if(teken.startsWith("BORD") || teken.startsWith("BCBORD")){
							System.out.println(pair.getKey() + "    teken: "+teken);
							
						}else {
							if(teken.equals("PYLA")) {
								//bewerk graden
								if(temp[22].contains("d")) {
									temp[22] = temp[22].substring(0, temp[22].indexOf("d"));
									}
								temp[22]= temp[22].replaceAll(".0000", "").replaceAll("g", "");
								
								
								try {
								int tmp = Integer.valueOf(temp[22])+135;
								if(tmp>=360) {
									tmp = tmp-360;
								}
								teken = "pijl " + tmp + " graden";
								}catch(Exception e) {
									try {
										int tmp = (int) (Double.valueOf(temp[22])+135);
										if(tmp>=360) {
											tmp = tmp-360;
										}
										teken = "pijl " + tmp + " graden";
									}catch(Exception e2) {
										System.out.println("error: pijl pad1: "+e.getMessage());
									}
								}
							}else if(teken.equals("PYLU")||teken.equals("PYLB")) {
								if(temp[22].contains("d")) {
								temp[22] = temp[22].substring(0, temp[22].indexOf("d"));
								}
								temp[22]= temp[22].replaceAll(".0000", "").replaceAll("g", "");
								
								try {
								int tmp = Integer.valueOf(temp[22])+180;
								if(tmp>=360) {
									tmp = tmp-360;
								}
								teken = "pijl " + tmp + " graden";
								}catch(Exception e) {
									try {
										int tmp = (int) (Double.valueOf(temp[22])+180);
										if(tmp>=360) {
											tmp = tmp-360;
										}
										teken = "pijl " + tmp + " graden";
									}catch(Exception e2) {
										System.out.println("error: pijl pad2: "+e.getMessage());
									}
								}
							}else if(teken.equals("nbpijl")) {
								if(temp[22].contains("d")) {
									temp[22] = temp[22].substring(0, temp[22].indexOf("d"));
									}
								temp[22]= temp[22].replaceAll(".0000", "").replaceAll("g", "");
								
								try {
								int tmp = Integer.valueOf(temp[22])+180;
								if(tmp>=360) {
									tmp = tmp-360;
								}
								teken = "pijl " + tmp + " graden";
								}catch(Exception e) {
									try {
										int tmp = (int) (Double.valueOf(temp[22])+180);
										if(tmp>=360) {
											tmp = tmp-360;
										}
										teken = "pijl " + tmp + " graden";
									}catch(Exception e2) {
										System.out.println("error: pijl pad3: "+e.getMessage());
									}
								}
							}else {
								if(temp[22].contains("d")) {
									temp[22] = temp[22].substring(0, temp[22].indexOf("d"));
								}
								temp[22]= temp[22].replaceAll(".0000", "").replaceAll("g", "");
								teken = "pijl " + temp[22] + " graden";
							}
						}
						if(teken.equals("")== false) {
						if (X > 63 && X < 108) {// linker rij bord
							if (Y > 162 && Y <= 175) { // 1
								voegToe(X, Y, teken, "pijl_bord_1_b",63,108,162,175);
							} else if (Y > 175 && Y < 188) { // 1
								voegToe(X, Y, teken, "pijl_bord_1_a",63,108,175,188);
							} else if (Y > 136 && Y <= 149.2) { // 3
								voegToe(X, Y, teken, "pijl_bord_3_b",63,108,136,149.2);
							} else if (Y > 149.2 && Y < 162) { // 3
								voegToe(X, Y, teken, "pijl_bord_3_a",63,108,149.2,162);
							} else if (Y > 110 && Y <= 122) { // 5
								voegToe(X, Y, teken, "pijl_bord_5_b",63,108,110,122);
							} else if (Y > 122 && Y < 135) { // 5
								voegToe(X, Y, teken, "pijl_bord_5_a",63,108,122,135);
							} else if (Y > 84 && Y < 97) { // 7
								voegToe(X, Y, teken, "pijl_bord_7_b",63,108,84,97);
							} else if (Y >= 97 && Y < 109) { // 7
								voegToe(X, Y, teken, "pijl_bord_7_a",63,108,97,109);
							} else if (Y > 58 && Y <= 70) { // 9
								voegToe(X, Y, teken, "pijl_bord_9_b",63,108,58,70);
							} else if (Y > 70 && Y < 83) { // 9
								voegToe(X, Y, teken, "pijl_bord_9_a",63,108,70,83);
							} else if (Y > 31 && Y <= 44) { // 11
								voegToe(X, Y, teken, "pijl_bord_11_b",63,108,31,44);
							} else if (Y > 44 && Y < 58) { // 11
								voegToe(X, Y, teken, "pijl_bord_11_a",63,108,44,58);
							}
						}else if (X/1000 > 63 && X/1000 < 108) {// linker rij bord
							if (Y/1000 > 162 && Y/1000 <= 175) { // 1
								voegToe(X/1000, Y/1000, teken, "pijl_bord_1_b",63,108,162,175);
							} else if (Y/1000 > 175 && Y/1000 < 188) { // 1
								voegToe(X/1000, Y/1000, teken, "pijl_bord_1_a",63,108,175,188);
							} else if (Y/1000 > 136 && Y/1000 <= 149.2) { // 3
								voegToe(X/1000, Y/1000, teken, "pijl_bord_3_b",63,108,136,149.2);
							} else if (Y/1000 > 149.2 && Y/1000 < 162) { // 3
								voegToe(X/1000, Y/1000, teken, "pijl_bord_3_a",63,108,149.2,162);
							} else if (Y/1000 > 110 && Y/1000 <= 122) { // 5
								voegToe(X/1000, Y/1000, teken, "pijl_bord_5_b",63,108,110,122);
							} else if (Y/1000 > 122 && Y/1000 < 135) { // 5
								voegToe(X/1000, Y/1000, teken, "pijl_bord_5_a",63,108,122,135);
							} else if (Y/1000 > 84 && Y/1000 < 97) { // 7
								voegToe(X/1000, Y/1000, teken, "pijl_bord_7_b",63,108,84,97);
							} else if (Y/1000 >= 97 && Y/1000 < 109) { // 7
								voegToe(X/1000, Y/1000, teken, "pijl_bord_7_a",63,108,97,109);
							} else if (Y/1000 > 58 && Y/1000 <= 70) { // 9
								voegToe(X/1000, Y/1000, teken, "pijl_bord_9_b",63,108,58,70);
							} else if (Y/1000 > 70 && Y/1000 < 83) { // 9
								voegToe(X/1000, Y/1000, teken, "pijl_bord_9_a",63,108,70,83);
							} else if (Y/1000 > 31 && Y/1000 <= 44) { // 11
								voegToe(X/1000, Y/1000, teken, "pijl_bord_11_b",63,108,31,44);
							} else if (Y/1000 > 44 && Y/1000 < 58) { // 11
								voegToe(X/1000, Y/1000, teken, "pijl_bord_11_a",63,108,44,58);
							}
						} else if (X > 110 && X < 155) {// rechter rij bord
							if (Y > 162 && Y < 175) { // 2
								voegToe(X, Y, teken, "pijl_bord_2_a",110,155,162,175);
							} else if (Y > 175 && Y < 188) { // 2
								voegToe(X, Y, teken, "pijl_bord_2_b",110,155,175,188);
							} else if (Y > 136 && Y <= 149.2) { // 4
								voegToe(X, Y, teken, "pijl_bord_4_a",110,155,136,149.2);
							} else if (Y > 149.2 && Y < 162) { // 4
								voegToe(X, Y, teken, "pijl_bord_4_b",110,155,149.2,162);
							} else if (Y > 110 && Y <= 122) { // 6
								voegToe(X, Y, teken, "pijl_bord_6_a",110,155,110,122);
							} else if (Y > 122 && Y < 135) { // 6
								voegToe(X, Y, teken, "pijl_bord_6_b",110,155,122,135);
							} else if (Y > 84 && Y < 97) { // 8
								voegToe(X, Y, teken, "pijl_bord_8_a",110,155,84,97);
							} else if (Y >= 97 && Y < 109) { // 8
								voegToe(X, Y, teken, "pijl_bord_8_b",110,155,97,109);
							} else if (Y > 58 && Y <= 70) { // 10
								voegToe(X, Y, teken, "pijl_bord_10_a",110,155,58,70);
							} else if (Y > 70 && Y < 83) { // 10
								voegToe(X, Y, teken, "pijl_bord_10_b",110,155,70,83);
							} else if (Y > 31 && Y <= 44) { // 12
								voegToe(X, Y, teken, "pijl_bord_12_a",110,155,31,44);
							} else if (Y > 44 && Y < 58) { // 12
								voegToe(X, Y, teken, "pijl_bord_12_b",110,155,44,58);
							}
						}else if (X/1000 > 110 && X/1000 < 155) {// rechter rij bord
							if (Y/1000 > 162 && Y/1000 < 175) { // 2
								voegToe(X/1000, Y/1000, teken, "pijl_bord_2_a",110,155,162,175);
							} else if (Y/1000 > 175 && Y/1000 < 188) { // 2
								voegToe(X/1000, Y/1000, teken, "pijl_bord_2_b",110,155,175,188);
							} else if (Y/1000 > 136 && Y/1000 <= 149.2) { // 4
								voegToe(X/1000, Y/1000, teken, "pijl_bord_4_a",110,155,136,149.2);
							} else if (Y/1000 > 149.2 && Y/1000 < 162) { // 4
								voegToe(X/1000, Y/1000, teken, "pijl_bord_4_b",110,155,149.2,162);
							} else if (Y/1000 > 110 && Y/1000 <= 122) { // 6
								voegToe(X/1000, Y/1000, teken, "pijl_bord_6_a",110,155,110,122);
							} else if (Y/1000 > 122 && Y/1000 < 135) { // 6
								voegToe(X/1000, Y/1000, teken, "pijl_bord_6_b",110,155,122,135);
							} else if (Y/1000 > 84 && Y/1000 < 97) { // 8
								voegToe(X/1000, Y/1000, teken, "pijl_bord_8_a",110,155,84,97);
							} else if (Y/1000 >= 97 && Y/1000 < 109) { // 8
								voegToe(X/1000, Y/1000, teken, "pijl_bord_8_b",110,155,97,109);
							} else if (Y/1000 > 58 && Y/1000 <= 70) { // 10
								voegToe(X/1000, Y/1000, teken, "pijl_bord_10_a",110,155,58,70);
							} else if (Y/1000 > 70 && Y/1000 < 83) { // 10
								voegToe(X/1000, Y/1000, teken, "pijl_bord_10_b",110,155,70,83);
							} else if (Y/1000 > 31 && Y/1000 <= 44) { // 12
								voegToe(X/1000, Y/1000, teken, "pijl_bord_12_a",110,155,31,44);
							} else if (Y/1000 > 44 && Y/1000 < 58) { // 12
								voegToe(X/1000, Y/1000, teken, "pijl_bord_12_b",110,155,44,58);
							}
						}
						}
					}
					else if (!teken.contains("bord") && !teken.contains("Attribute Definition")&& !teken.equals("p")) {
						if (teken.endsWith("A$C4BB11CC5")) {
							teken = "KERK";
						}
						if (X > 63 && X < 108) {// linker rij bord
							if (Y > 162 && Y <= 175) { // 1
								voegToe(X, Y, teken, "teken_bord_1_b",63,108,162,175);
							} else if (Y > 175 && Y < 188) { // 1
								voegToe(X, Y, teken, "teken_bord_1_a",63,108,175,188);
							} else if (Y > 136 && Y <= 149.2) { // 3
								voegToe(X, Y, teken, "teken_bord_3_b",63,108,136,149.2);
							} else if (Y > 149.2 && Y < 162) { // 3
								voegToe(X, Y, teken, "teken_bord_3_a",63,108,149.2,162);
							} else if (Y > 110 && Y <= 122) { // 5
								voegToe(X, Y, teken, "teken_bord_5_b",63,108,110,122);
							} else if (Y > 122 && Y < 135) { // 5
								voegToe(X, Y, teken, "teken_bord_5_a",63,108,122,135);
							} else if (Y > 84 && Y < 97) { // 7
								voegToe(X, Y, teken, "teken_bord_7_b",63,108,84,97);
							} else if (Y >= 97 && Y < 109) { // 7
								voegToe(X, Y, teken, "teken_bord_7_a",63,108,97,109);
							} else if (Y > 58 && Y <= 70) { // 9
								voegToe(X, Y, teken, "teken_bord_9_b",63,108,58,70);
							} else if (Y > 70 && Y < 83) { // 9
								voegToe(X, Y, teken, "teken_bord_9_a",63,108,70,83);
							} else if (Y > 31 && Y <= 44) { // 11
								voegToe(X, Y, teken, "teken_bord_11_b",63,108,31,44);
							} else if (Y > 44 && Y < 58) { // 11
								voegToe(X, Y, teken, "teken_bord_11_a",63,108,44,58);
							}
						
						}else if (X/1000 > 63 && X/1000 < 108) {// linker rij bord
							if (Y/1000 > 162 && Y/1000 <= 175) { // 1
								voegToe(X/1000, Y/1000, teken, "teken_bord_1_b",63,108,162,175);
							} else if (Y/1000 > 175 && Y/1000 < 188) { // 1
								voegToe(X/1000, Y/1000, teken, "teken_bord_1_a",63,108,175,188);
							} else if (Y/1000 > 136 && Y/1000 <= 149.2) { // 3
								voegToe(X/1000, Y/1000, teken, "teken_bord_3_b",63,108,136,149.2);
							} else if (Y/1000 > 149.2 && Y/1000 < 162) { // 3
								voegToe(X/1000, Y/1000, teken, "teken_bord_3_a",63,108,149.2,162);
							} else if (Y/1000 > 110 && Y/1000 <= 122) { // 5
								voegToe(X/1000, Y/1000, teken, "teken_bord_5_b",63,108,110,122);
							} else if (Y/1000 > 122 && Y/1000 < 135) { // 5
								voegToe(X/1000, Y/1000, teken, "teken_bord_5_a",63,108,122,135);
							} else if (Y/1000 > 84 && Y/1000 < 97) { // 7
								voegToe(X/1000, Y/1000, teken, "teken_bord_7_b",63,108,84,97);
							} else if (Y/1000 >= 97 && Y/1000 < 109) { // 7
								voegToe(X/1000, Y/1000, teken, "teken_bord_7_a",63,108,97,109);
							} else if (Y/1000 > 58 && Y/1000 <= 70) { // 9
								voegToe(X/1000, Y/1000, teken, "teken_bord_9_b",63,108,58,70);
							} else if (Y/1000 > 70 && Y/1000 < 83) { // 9
								voegToe(X/1000, Y/1000, teken, "teken_bord_9_a",63,108,70,83);
							} else if (Y/1000 > 31 && Y/1000 <= 44) { // 11
								voegToe(X/1000, Y/1000, teken, "teken_bord_11_b",63,108,31,44);
							} else if (Y/1000 > 44 && Y/1000 < 58) { // 11
								voegToe(X/1000, Y/1000, teken, "teken_bord_11_a",63,108,44,58);
							}
						
						} else if (X > 110 && X < 155) {// rechter rij bord
							if (Y > 162 && Y < 175) { // 2
								voegToe(X, Y, teken, "teken_bord_2_a",110,155,162,175);
							} else if (Y > 175 && Y < 188) { // 2
								voegToe(X, Y, teken, "teken_bord_2_b",110,155,175,188);
							} else if (Y > 136 && Y <= 149.2) { // 4
								voegToe(X, Y, teken, "teken_bord_4_a",110,155,136,149.2);
							} else if (Y > 149.2 && Y < 162) { // 4
								voegToe(X, Y, teken, "teken_bord_4_b",110,155,149.2,162);
							} else if (Y > 110 && Y <= 122) { // 6
								voegToe(X, Y, teken, "teken_bord_6_a",110,155,110,122);
							} else if (Y > 122 && Y < 135) { // 6
								voegToe(X, Y, teken, "teken_bord_6_b",110,155,122,135);
							} else if (Y > 84 && Y < 97) { // 8
								voegToe(X, Y, teken, "teken_bord_8_a",110,155,84,97);
							} else if (Y >= 97 && Y < 109) { // 8
								voegToe(X, Y, teken, "teken_bord_8_b",110,155,97,109);
							} else if (Y > 58 && Y <= 70) { // 10
								voegToe(X, Y, teken, "teken_bord_10_a",110,155,58,70);
							} else if (Y > 70 && Y < 83) { // 10
								voegToe(X, Y, teken, "teken_bord_10_b",110,155,70,83);
							} else if (Y > 31 && Y <= 44) { // 12
								voegToe(X, Y, teken, "teken_bord_12_a",110,155,31,44);
							} else if (Y > 44 && Y < 58) { // 12
								voegToe(X, Y, teken, "teken_bord_12_b",110,155,44,58);
							}
						}else if (X/1000 > 110 && X/1000 < 155) {// rechter rij bord
							if (Y/1000 > 162 && Y/1000 < 175) { // 2
								voegToe(X/1000, Y/1000, teken, "teken_bord_2_a",110,155,162,175);
							} else if (Y/1000 > 175 && Y/1000 < 188) { // 2
								voegToe(X/1000, Y/1000, teken, "teken_bord_2_b",110,155,175,188);
							} else if (Y/1000 > 136 && Y/1000 <= 149.2) { // 4
								voegToe(X/1000, Y/1000, teken, "teken_bord_4_a",110,155,136,149.2);
							} else if (Y/1000 > 149.2 && Y/1000 < 162) { // 4
								voegToe(X/1000, Y/1000, teken, "teken_bord_4_b",110,155,149.2,162);
							} else if (Y/1000 > 110 && Y/1000 <= 122) { // 6
								voegToe(X/1000, Y/1000, teken, "teken_bord_6_a",110,155,110,122);
							} else if (Y/1000 > 122 && Y/1000 < 135) { // 6
								voegToe(X/1000, Y/1000, teken, "teken_bord_6_b",110,155,122,135);
							} else if (Y/1000 > 84 && Y/1000 < 97) { // 8
								voegToe(X/1000, Y/1000, teken, "teken_bord_8_a",110,155,84,97);
							} else if (Y/1000 >= 97 && Y/1000 < 109) { // 8
								voegToe(X/1000, Y/1000, teken, "teken_bord_8_b",110,155,97,109);
							} else if (Y/1000 > 58 && Y/1000 <= 70) { // 10
								voegToe(X/1000, Y/1000, teken, "teken_bord_10_a",110,155,58,70);
							} else if (Y/1000 > 70 && Y/1000 < 83) { // 10
								voegToe(X/1000, Y/1000, teken, "teken_bord_10_b",110,155,70,83);
							} else if (Y/1000 > 31 && Y/1000 <= 44) { // 12
								voegToe(X/1000, Y/1000, teken, "teken_bord_12_a",110,155,31,44);
							} else if (Y/1000 > 44 && Y/1000 < 58) { // 12
								voegToe(X/1000, Y/1000, teken, "teken_bord_12_b",110,155,44,58);
							}
						}
					}
					
					}catch(Exception e) {
						System.out.println("test pijlen en tekens: "+e.getMessage());
					}
					}
				}

			}
			
			if(bordSpec==5 && borddata==false) {
				ArrayList<Object[]> tempList = verzameling.get("borddata");
				try {
				if(tempList != null) {
					
				double highestY=-999999;
				double highestY2 = -999999;
				double highestY3 = -999999;
				for(int i=0;i<tempList.size();i++) {
					if((Double) tempList.get(i)[0]>highestY) {
						highestY = (Double) tempList.get(i)[0];
					}
				}
				
				for(int i=0;i<tempList.size();i++) {
					if((Double) tempList.get(i)[0]<highestY &&(Double) tempList.get(i)[0]>highestY2 ) {
						highestY2 = (Double) tempList.get(i)[0];
					}
				}
				for(int i=0;i<tempList.size();i++) {
					if((Double) tempList.get(i)[0]<highestY2 &&(Double) tempList.get(i)[0]>highestY3 ) {
						highestY3 = (Double) tempList.get(i)[0];
					}
				}
				ArrayList<Object[]> BxHtemp = new ArrayList<Object[]>();
				ArrayList<Object[]> LTypetemp = new ArrayList<Object[]>();
				for(int i=0;i<tempList.size();i++) {
					if((Double) tempList.get(i)[0]<highestY3) {
						voegToe((Double) tempList.get(i)[0], (Double) tempList.get(i)[1], (String) tempList.get(i)[2], "Omschrijving");
					}else if((Double) tempList.get(i)[0]==highestY3) {
						String stand = (String) tempList.get(i)[2];
						if(stand.startsWith("N")|| stand.startsWith("A"))
						voegToe((Double) tempList.get(i)[0], (Double) tempList.get(i)[1], (String) tempList.get(i)[2], "standplaats");
					}else if((Double) tempList.get(i)[0]==highestY2) {
						Object[] temp = new Object[2];
						temp[0] = (Double) tempList.get(i)[1];
						temp[1] = (String) tempList.get(i)[2];
						BxHtemp.add(temp);
					}else if((Double) tempList.get(i)[0]==highestY) {
						Object[] temp = new Object[2];
						temp[0] = (Double) tempList.get(i)[1];
						temp[1] = (String) tempList.get(i)[2];
						LTypetemp.add(temp);
					}
				}
				
				if(BxHtemp.size()>1) {
					boolean klaar =false;
					while(klaar == false) {
						klaar= true;
						for(int i=0;i<BxHtemp.size()-1;i++) {
							if((Double) BxHtemp.get(i)[0]>(Double) BxHtemp.get(i+1)[0]) {
								Collections.swap(BxHtemp, i, i+1);
								klaar = false;
							}
						}
					}
				}
				
				if(LTypetemp.size()>1) {
					boolean klaar =false;
					while(klaar == false) {
						klaar= true;
						for(int i=0;i<LTypetemp.size()-1;i++) {
							if((Double) LTypetemp.get(i)[0]>(Double) LTypetemp.get(i+1)[0]) {
								Collections.swap(LTypetemp, i, i+1);
								klaar = false;
							}
						}
					}
				}
				if(BxHtemp.size()==1) {
					String BxH = (String) BxHtemp.get(0)[1];
					if(BxH.contains("x")) {
						String[] tmp = BxH.split("x");
						results[30] =tmp[0]; //BW
						results[31] =tmp[1]; //BH
					}
				}else if(BxHtemp.size()==2) {
					results[30] =(String) BxHtemp.get(0)[1]; //BW
					results[31] =(String) BxHtemp.get(1)[1]; //BH
				}else if(BxHtemp.size()==3) {
					results[30] =(String) BxHtemp.get(0)[1]; //BW
					results[31] =(String) BxHtemp.get(2)[1]; //BH
				}
				
				if(LTypetemp.size()>0) {
					String Ltype = (String) LTypetemp.get(0)[1];
					if(Ltype.contains("/")) {
						String[] tmp = Ltype.split("/");
						results[18] =tmp[0]; //uppercase
						results[19] =tmp[1]; //lowercase
						if(LTypetemp.size()>1) {
							results[20] = (String) LTypetemp.get(1)[1];
						}
					}
					results[18]=Ltype;
					if(LTypetemp.size()>1) {
						results[19] = (String) LTypetemp.get(1)[1];
					}
					if(LTypetemp.size()>2) {
						results[20] = (String) LTypetemp.get(2)[1];
					}
				}
				}
				}catch(Exception e) {
					System.out.println("test 10: "+e.getMessage());
				}
			}
			
			if(results[17] == null || results[17].isBlank()) {
			results[17] = stand2;
			}else {			
					results[17] = results[17]+" "+stand2;				
			}
			
			sorteer();
			

			printResultaat(pair.getKey());
		}

	}

	



@SuppressWarnings({ "rawtypes", "unchecked" })
private static void printObjecten() {
		
		
		sheet = workbook.getSheetAt(5); // Objecten dubbel

		int count=1;
		
		
		Iterator it =objectenTotaal.entrySet().iterator();
		System.out.println(objectenTotaal.size());
		int counter =0;
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy", Locale.ENGLISH);
		DateTimeFormatter formatter2 = DateTimeFormatter.ofPattern("dd-M-yyyy", Locale.ENGLISH);
		while (it.hasNext()) {
			counter++;
			
			if(counter % 100 ==0) {
				System.out.println("bestand nr: "+counter);
			}
			HashMap.Entry<String, ArrayList<String[]>> pair = (HashMap.Entry) it.next();
			
			
			ArrayList<String[]> Lijst = pair.getValue();
			
			if(pair.getKey() == null || pair.getKey().isBlank()) {
				
			}else {
			if(Lijst.size()>1) {
				//sorteren nieuwste is .get(0)
				System.out.println("dubbele bestanden: "+Lijst.size());
				for(int x=0;x<30;x++) {
					boolean changed = false;
					for(int i=0;i<Lijst.size()-1;i++) {
						
						String[] tmp = Lijst.get(i);
						String[] tmp2 = Lijst.get(i+1);
						
						try {
							LocalDate dateTime;
						try {
						dateTime = LocalDate.parse(tmp[3], formatter);
						}catch(DateTimeParseException e) {
							try {
							dateTime = LocalDate.parse(tmp[3], formatter2);
							}catch(DateTimeParseException e2) {
								dateTime = null;
							}
						}
						LocalDate dateTime2;
						try {
							dateTime2 = LocalDate.parse(tmp2[3], formatter);
							}catch(DateTimeParseException e) {
								try {
								dateTime2 = LocalDate.parse(tmp2[3], formatter2);
								}catch(DateTimeParseException e2) {
									dateTime2 = null;
								}
							}
						if(dateTime!= null && dateTime2!= null && dateTime.isBefore(dateTime2)) {
							//switch
							Collections.swap(Lijst, i, i+1);
							changed = true;
							break;
						}
						}catch(NullPointerException e){
							Collections.swap(Lijst, i, i+1);
							break;
						}
					}
					
					if(changed == false) {
						break;
					}
				}
				
			}
			//System.gc();
			for(int i=0;i<Lijst.size();i++) {
			row = sheet.getRow(count);

			if (row == null) {
				row = sheet.createRow(count);
			}
			String[] tmp = Lijst.get(i);
			if(pair.getValue().size()==1) {
				tmp[2] = "Nee"; //dubbel
				tmp[4] = (i+1)+"";   //rang
			}else {
				tmp[2] = "Ja";
				tmp[4] = (i+1)+"";   //rang
			}
			
			for(int x=0;x<tmp.length;x++) {
				cell = row.getCell(x);
				if (cell == null) {
					cell = row.createCell(x);
				}
				cell.setCellValue(tmp[x]);
			}
			
			count++;
			}
			}
		}
		System.gc();
	}

	private static void printResultaat(String bestand) {
		try {
		if (doelLocatie.exists()) {
			for(int i=0;i<results.length;i++) {
				if(results[i]==null) {
					results[i]="";
				}
			}
			try {
			sheet = workbook.getSheetAt(0); // Tekenhoofd

			row = sheet.getRow(SpecsheetRow);

			if (row == null) {
				row = sheet.createRow(SpecsheetRow);
			}

			for (int i = 0; i < 7; i++) {
				if (results[i] != null && !results[i].equals("")) {
					cell = row.getCell(i);
					if (cell == null) {
						cell = row.createCell(i);
					}
					if(i==3) {
						if(results[i]!= null &&results[i].equals("")==false) {
							String[] temp = results[i].split("-");
							if(temp.length==3) {
								String bouwjaar = (String) temp[2];
								if(bouwjaar.length()==2) {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
									}else {
										bouwjaar ="19"+jaar;
									}
									results[i] = temp[0]+"-"+temp[1]+"-"+bouwjaar;
								}
							}
						}
					}
					cell.setCellValue(results[i]);
				}
			}
			cell = row.getCell(7);
			if (cell == null) {
				cell = row.createCell(7);
			}
			
			cell.setCellValue(bordSpec);
			
			ArrayList<String[]> tmpList = objectenTotaal.get(results[6]);
			if(tmpList ==null) {
				tmpList = new ArrayList<String[]>();
			}
			String[] tmpArray = new String[5];
			tmpArray[0] = results[1];
			tmpArray[1] = results[6];
			tmpArray[3] = results[3];
			
			tmpList.add(tmpArray);
			
			objectenTotaal.put(results[6], tmpList);
			
			SpecsheetRow++;
			}catch(Exception e) {
				System.out.println("catch tab 0: "+e.getMessage());
			}
			///////////////////////////////////////////////////////////////////
			try {
			sheet = workbook.getSheetAt(1); // Object

			row = sheet.getRow(ObjectRow);

			if (row == null) {
				row = sheet.createRow(ObjectRow);
			}

			ArrayList<Object[]> tempListStandplaats = verzameling.get("standplaats");
			if(tempListStandplaats != null) {
				if(tempListStandplaats.size()==1) {
					results[17] = (String) tempListStandplaats.get(0)[2];
				}
				if(tempListStandplaats.size()>1) {
					results[17] = (String) tempListStandplaats.get(0)[2];
					for(int i=1;i<tempListStandplaats.size();i++) {
						results[17] = results[17]+" "+(String) tempListStandplaats.get(i)[2];
					}
				}
			}
			results[17] = results[17].replace("Standplaats:", "");
			results[17] = results[17].strip();
			
			String[] tempSplit = results[17].split(" ");
			for(int i=0;i<tempSplit.length;i++) {
				if(tempSplit[i].contains("-")&&tempSplit[i].contains(",") &&!tempSplit[i].contains("'")&&!tempSplit[i].contains("`")&&!tempSplit[i].contains("(")&&!tempSplit[i].contains(")")&&!tempSplit[i].contains("/") ) {
					if(tempSplit[i].substring(tempSplit[i].length()-1).contains(",")||tempSplit[i].substring(tempSplit[i].length()-1).contains("m")) {
						
					}else {
					BPScodering = tempSplit[i];
					break;
					}
				}
			}
			int counter = 1;

			cell = row.getCell(0);
			if (cell == null) {
				cell = row.createCell(0);
			}
			
			cell.setCellValue(results[1]); // bestandsnaam

			for (int i = 7; i < 18; i++) {
				if (results[i] != null && !results[i].isEmpty()) {
					cell = row.getCell(counter);
					if (cell == null) {
						cell = row.createCell(counter);
					}
					cell.setCellValue(results[i]);

				}
				counter++;
			}
			
			
			
			String consulkring="";
			ArrayList<Object[]> tempListConsulkring = verzameling.get("Consulkring");
			if(tempListConsulkring != null) {
				if(tempListConsulkring.size()==1) {
					consulkring = (String) tempListConsulkring.get(0)[2];
				}
				if(tempListConsulkring.size()>1) {
					consulkring = (String) tempListConsulkring.get(0)[2];
					for(int i=1;i<tempListConsulkring.size();i++) {
						consulkring = consulkring+" "+(String) tempListConsulkring.get(i)[2];
					}
				}
			}
			
			consulkring = consulkring.replace("Consulkring:", "");
			consulkring = consulkring.strip();
			if(!consulkring.isEmpty()) {
				cell = row.getCell(12);
				if (cell == null) {
					cell = row.createCell(12);
				}
			cell.setCellValue(consulkring);
			}
			
			
			String opmerkingen="";
			ArrayList<Object[]> tempListOpmerkingen = verzameling.get("Opmerkingen");
			if(tempListOpmerkingen != null) {
				if(tempListOpmerkingen.size()==1) {
					opmerkingen = (String) tempListOpmerkingen.get(0)[2];
				}
				if(tempListOpmerkingen.size()>1) {
					opmerkingen = (String) tempListOpmerkingen.get(0)[2];
					for(int i=1;i<tempListOpmerkingen.size();i++) {
						opmerkingen = opmerkingen+" "+(String) tempListOpmerkingen.get(i)[2];
					}
				}
			}
			opmerkingen = opmerkingen.replace("Opmerkingen", "");
			opmerkingen = opmerkingen.strip();
			if(!opmerkingen.isEmpty()) {
				cell = row.getCell(13);
				if (cell == null) {
					cell = row.createCell(13);
				}
				cell.setCellValue(opmerkingen);
			}
			
			
			if(BPScodering == null || BPScodering.isEmpty()) {
				tempSplit = opmerkingen.split(" ");
				for(int i=0;i<tempSplit.length;i++) {
					if(tempSplit[i].contains("-")&&tempSplit[i].contains(",") &&!tempSplit[i].contains("'")&&!tempSplit[i].contains("`")&&!tempSplit[i].contains("(")&&!tempSplit[i].contains(")")&&!tempSplit[i].contains("/") ) {
						if(tempSplit[i].substring(tempSplit[i].length()-1).contains(",")||tempSplit[i].substring(tempSplit[i].length()-1).contains("m")) {
						
						}else {
						BPScodering = tempSplit[i];
						break;
						}
					}
				}
			}
			
			if(BPScodering!= null && !BPScodering.isEmpty()) {
			cell = row.getCell(14);
			if (cell == null) {
				cell = row.createCell(14);
			}

			cell.setCellValue(BPScodering);
			}
			ObjectRow++;
			}catch(Exception e) {
				System.out.println("catch tab 1: "+e.getMessage());
			}
			/////////////////////////////////////////////////////////////////////////////
			try {
			sheet = workbook.getSheetAt(2); // Armgegevens

			if (true) {//bordSpec == 1
				for (int i = 1; i < 13; i++) {
					String keyA = "woord_bord_" + i + "_a";
					String keyB = "woord_bord_" + i + "_b";

					String keyA2 = "teken_bord_" + i + "_a";
					String keyB2 = "teken_bord_" + i + "_b";
					
					String keyA3 = "pijl_bord_" + i + "_a";
					String keyB3 = "pijl_bord_" + i + "_b";

					ArrayList<String> tempList = verzameling2.get(keyA);
					ArrayList<Object[]> tempList2 = verzameling.get(keyA);

					ArrayList<String> tempList3 = verzameling2.get(keyB);
					ArrayList<Object[]> tempList4 = verzameling.get(keyB);

					ArrayList<Object[]> tempListTekenA = verzameling.get(keyA2);
					ArrayList<Object[]> tempListTekenB = verzameling.get(keyB2);
					
					ArrayList<Object[]> tempListPijlA = verzameling.get(keyA3);
					ArrayList<Object[]> tempListPijlB = verzameling.get(keyB3);
					
					ArrayList<Object[]> tempListOnbekend = verzameling.get("pijl_bord_onbekend");
					String onbekend ="";
					if(tempListOnbekend != null && tempListOnbekend.size()>0) {
						
						onbekend = (String) tempListOnbekend.get(0)[2]; // onbekend object (pijl)
					}
					
					if (tempList != null || tempList2 != null || tempList3 != null || tempList4 != null) {
						String[] result = new String[53];
						try {
						result[0] = results[7]; // object
						result[1] = i + ""; // arm
						result[2] = results[18];
						result[3] = results[19];
						result[4] = results[20];
						result[5] = results[21];
						result[6] = results[22];//kleur
						result[7] =	results[30];//BW
						result[8] = results[31];//BH
						
						result[37] = onbekend; // onbekende objecten
						}catch(Exception e) {
							System.out.println("catch test 1");
						}
						try {
						ArrayList<String> tempListA = new ArrayList<String>();
						if (tempList != null) {
							
							for (int x = 0; x < tempList.size(); x++) {
								if(tempList.get(x).equals(BPScodering)||tempList.get(x).equals(results[7])) {
									
								}else {
								tempListA.add(tempList.get(x));
								}
							}
							
						}
						if (tempList2 != null) {
							for (int x = 0; x < tempList2.size(); x++) {
								if(tempList2.get(x)[2].equals(BPScodering)||tempList2.get(x)[2].equals(results[7])) {
									
								}else {
								tempListA.add((String) tempList2.get(x)[2]);
								}
							}
							
						}
						
						ArrayList<String> tempListB = new ArrayList<String>();
						if (tempList3 != null) {
							for (int x = 0; x < tempList3.size(); x++) {
								if(tempList3.get(x).equals(BPScodering)||tempList3.get(x).equals(results[7])) {
									
								}else {
								tempListB.add(tempList3.get(x));
								}
							}
						}
						if (tempList4 != null) {
							for (int x = 0; x < tempList4.size(); x++) {
								if(tempList4.get(x)[2].equals(BPScodering)||tempList4.get(x)[2].equals(results[7])) {
									
								}else {
								tempListB.add((String) tempList4.get(x)[2]);
								}
							}
						}

						
						if (tempListA.size() > 0) {
							
							String[] tmp = tempListA.get(0).split(" ");
							
		
							if(tmp.length ==1) {
								result[9] = tmp[0];
							}else if(tmp.length ==2) {							
								if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
									result[9] = tempListA.get(0);
								}else {
									if(isNumeric(tmp[1])) {
										result[9] = tmp[0];
										
										result[10] = tmp[1];
										
									}else {
										result[9] = tempListA.get(0);
									}
								

								}
							}else {
								
								if(tmp[0]!= null) {
									if(isNumeric(tmp[tmp.length-1])&&
											(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[9] = tot;
									result[10] = tmp[tmp.length-1];
									}else {
										result[9] = tempListA.get(0);
									}
								}
							}
							
							
							if (tempListA.size() > 1) {
								tmp = tempListA.get(1).split(" ");
								
								if(tmp.length ==1) {
									result[11] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[11] = tempListA.get(1);
									}else {
										if(isNumeric(tmp[1])) {
											result[11] = tmp[0];
											
											result[12] = tmp[1];
										}else {
											result[11] = tempListA.get(1);
										}
								

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[11] = tot;
									result[12] = tmp[tmp.length-1];
										}else {
											result[11] = tempListA.get(1);
										}
									}
								}

							}
							
							if (tempListA.size() > 2) {
								tmp = tempListA.get(2).split(" ");
								
								if(tmp.length ==1) {
									result[13] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[13] = tempListA.get(2);
									}else {
										if(isNumeric(tmp[1])) {
											result[13] = tmp[0];
											
											result[14] = tmp[1];
										}else {
											result[13] = tempListA.get(2);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[13] = tot;
									result[14] = tmp[tmp.length-1];
										}else {
											result[13] = tempListA.get(2);
										}
									}
								}
						
							}

							if (tempListA.size() > 3) {
								tmp = tempListA.get(3).split(" ");
								
								if(tmp.length ==1) {
									result[15] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[15] = tempListA.get(3);
									}else {
										if(isNumeric(tmp[1])) {
											result[15] = tmp[0];
											
											result[16] = tmp[1];
										}else {
											result[15] = tempListA.get(3);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[15] = tot;
									result[16] = tmp[tmp.length-1];
										}else {
											result[15] = tempListA.get(3);
										}
									}
								}
														
							}
							if (tempListA.size() > 4) {
								tmp = tempListA.get(4).split(" ");
								
								if(tmp.length ==1) {
									result[17] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[17] = tempListA.get(4);
									}else {
										if(isNumeric(tmp[1])) {
											result[17] = tmp[0];
											
											result[18] = tmp[1];
										}else {
											result[17] = tempListA.get(4);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[17] = tot;
									result[18] = tmp[tmp.length-1];
										}else {
											result[17] = tempListA.get(4);
										}
									}
								}					
							}
							if (tempListA.size() > 5) {
								tmp = tempListA.get(5).split(" ");
								
								if(tmp.length ==1) {
									result[19] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[19] = tempListA.get(5);
									}else {
										if(isNumeric(tmp[1])) {
											result[19] = tmp[0];
											
											result[20] = tmp[1];
										}else {
											result[19] = tempListA.get(5);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[19] = tot;
									result[20] = tmp[tmp.length-1];
										}else {
											result[19] = tempListA.get(5);
										}
									}
								}			
							}
						}

						if (tempListTekenA != null && tempListTekenA.size() > 0) {
							String text = "";
							
							for (int x = 0; x < tempListTekenA.size(); x++) {
								text = text + tempListTekenA.get(x)[2];
								if (x < tempListTekenA.size() - 1) {
									text = text + " | ";
								}
							}
							result[22] = text; // speciale tekens A
							
							
						}
						
						if (tempListPijlA != null && tempListPijlA.size() > 0) {
							String text = "";
							
							for (int x = 0; x < tempListPijlA.size(); x++) {
								
								text = text + tempListPijlA.get(x)[2];
								if (x < tempListPijlA.size() - 1) {
									text = text + " | ";
								}
							}
							
							result[21] = text; //Pijlen A
							
						}

						if (tempListB.size() > 0) {
							String[] tmp = tempListB.get(0).split(" ");
							
							if(tmp.length ==1) {
								result[23] = tmp[0];
							}else if(tmp.length ==2) {							
								if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
									result[23] = tempListB.get(0);
								}else {
									if(isNumeric(tmp[1])) {
										result[23] = tmp[0];
										
										result[24] = tmp[1];
									}else {
										result[23] = tempListB.get(0);
									}
								

								}
							}else {
								
								if(tmp[0]!= null) {
									if(isNumeric(tmp[tmp.length-1])&&
											(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
								String tot = tmp[0];
								for(int x=1;x<tmp.length-1;x++) {
									tot = tot+" "+tmp[x];
								}
								result[23] = tot;
								result[24] = tmp[tmp.length-1];
									}else {
										result[23] = tempListB.get(0);
									}
								}
							}

							if (tempListB.size() > 1) {
								tmp = tempListB.get(1).split(" ");
								
								if(tmp.length ==1) {
									result[25] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[25] = tempListB.get(1);
									}else {
										if(isNumeric(tmp[1])) {
											result[25] = tmp[0];
											
											result[26] = tmp[1];
										}else {
											result[25] = tempListB.get(1);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[25] = tot;
									result[26] = tmp[tmp.length-1];
										}else {
											result[25] = tempListB.get(1);
										}
									}
								}								
							}
							
							if (tempListB.size() > 2) {
								tmp = tempListB.get(2).split(" ");
								
								if(tmp.length ==1) {
									result[27] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[27] = tempListB.get(2);
									}else {
										if(isNumeric(tmp[1])) {
											result[27] = tmp[0];
											
											result[28] = tmp[1];
										}else {
											result[27] = tempListB.get(2);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[27] = tot;
									result[28] = tmp[tmp.length-1];
										}else {
											result[27] = tempListB.get(2);
										}
									}
								}						
							}
							if (tempListB.size() > 3) {
								tmp = tempListB.get(3).split(" ");
								
								if(tmp.length ==1) {
									result[29] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[29] = tempListB.get(3);
									}else {
										if(isNumeric(tmp[1])) {
											result[29] = tmp[0];
											
											result[30] = tmp[1];
										}else {
											result[29] = tempListB.get(3);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[29] = tot;
									result[30] = tmp[tmp.length-1];
										}else {
											result[29] = tempListB.get(3);
										}
									}
								}					
							}
							if (tempListB.size() > 4) {
								tmp = tempListB.get(4).split(" ");
								
								if(tmp.length ==1) {
									result[31] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[31] = tempListB.get(4);
									}else {
										if(isNumeric(tmp[1])) {
											result[31] = tmp[0];
											
											result[32] = tmp[1];
										}else {
											result[31] = tempListB.get(4);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[31] = tot;
									result[32] = tmp[tmp.length-1];
										}else {
											result[31] = tempListB.get(4);
										}
									}
								}						
							}
							if (tempListB.size() > 5) {
								tmp = tempListB.get(5).split(" ");
								
								if(tmp.length ==1) {
									result[33] = tmp[0];
								}else if(tmp.length ==2) {							
									if(tmp[0].contentEquals("N")||tmp[0].contentEquals("A")||tmp[0].contentEquals("S")) {
										result[33] = tempListB.get(5);
									}else {
										if(isNumeric(tmp[1])) {
											result[33] = tmp[0];
											
											result[34] = tmp[1];
										}else {
											result[33] = tempListB.get(5);
										}
									

									}
								}else {
									
									if(tmp[0]!= null) {
										if(isNumeric(tmp[tmp.length-1])&&
												(tmp.length<2 || !(tmp[tmp.length-2].contentEquals("N")|| tmp[tmp.length-2].contentEquals("A")|| tmp[tmp.length-2].contentEquals("S")))) {
									String tot = tmp[0];
									for(int x=1;x<tmp.length-1;x++) {
										tot = tot+" "+tmp[x];
									}
									result[33] = tot;
									result[34] = tmp[tmp.length-1];
										}else {
											result[33] = tempListB.get(5);
										}
									}
								}						
							}
						}
						if (tempListTekenB != null && tempListTekenB.size() > 0) {
							String text = "";
							for (int x = 0; x < tempListTekenB.size(); x++) {
								text = text + tempListTekenB.get(x)[2];
								if (x < tempListTekenB.size() - 1) {
									text = text + " | ";
								}
							}
							result[36] = text; // speciale tekens B
							
						}

						if (tempListPijlB != null && tempListPijlB.size() > 0) {
							String text = "";
							
							for (int x = 0; x < tempListPijlB.size(); x++) {
								text = text + tempListPijlB.get(x)[2];
								if (x < tempListPijlB.size() - 1) {
									text = text + " | ";
								}
							}
							
							result[35] = text; //Pijlen B
							
						}
						}catch(Exception e) {
							System.out.println("catch test 2");
						}
						try {
						if (klokstanden.get(String.valueOf(i)) != null) {
							
							result[38] = klokstanden.get(String.valueOf(i))[0];
							if(result[38].endsWith("0")||result[38].endsWith("5")) {
								
							}else {
								//geen meervoud van 5
								if(result[38].endsWith("1")||result[38].endsWith("2")) {
									
									result[38] = result[38].substring(0, result[38].length()-1) +"0";
									
								}else if(result[38].endsWith("8")||result[38].endsWith("9")) {
									
									result[38] = result[38].substring(0, result[38].length()-1) +"0";
									result[38] = (Integer.valueOf(result[38])+10)+"";
									
									
								}else if(result[38].endsWith("3")||result[38].endsWith("4")
										||result[38].endsWith("6")||result[38].endsWith("7")) {
									
									result[38] = result[38].substring(0, result[38].length()-1) +"0";
									
								}
								
							}
						}
						}catch(Exception e) {
							System.out.println("catch test 3");
						}
						
						keyA = "TAG_" + i;
						
						ArrayList<Object[]> ListA = verzameling.get(keyA);
						
						if (ListA != null) {
							if (ListA.size() == 1) {
								result[39] = "";
								result[40] = "";
								result[41] = (String) ListA.get(0)[2];
							} else if (ListA.size() == 2) {
								result[39] = (String) ListA.get(0)[2];
								result[40] = (String) ListA.get(1)[2];
								result[41] = "";
							} else if (ListA.size() > 2) {
								result[39] = (String) ListA.get(0)[2];
								result[40] = (String) ListA.get(2)[2];
								result[41] = (String) ListA.get(1)[2];
							}

						}
						try {
						keyA = "bj_" + i;

						ListA = verzameling.get(keyA);

						if (ListA != null) {
							if (ListA.size() == 1) {
								result[42] = "";
								result[43] = "";
								String bouwjaar = ((String) ListA.get(0)[2]).replaceAll("!","1");
								
								if(bouwjaar.length()==2) {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
								}
								result[44] = bouwjaar;
							} else if (ListA.size() == 2 && ListA.get(0)[2]!=null) {
								String bouwjaar = ((String) ListA.get(0)[2]).replaceAll("!","1");
								if(bouwjaar.length()==2) {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
								}
								result[42] = bouwjaar;
								bouwjaar = ((String) ListA.get(1)[2]).replaceAll("!","1");
								if(bouwjaar.length()==2) {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
								}
								result[43] = bouwjaar;								
								result[44] = "";
							} else if (ListA.size() > 2 && ListA.get(0)[2]!=null) {
								String bouwjaar = ((String) ListA.get(0)[2]).replaceAll("!","1");
								if(bouwjaar.length()==2 ) {
									try {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
									}catch(NumberFormatException e) {
										
									}
								}
								result[42] = bouwjaar;
								
								bouwjaar = ((String) ListA.get(2)[2]).replaceAll("!","1");
								if(bouwjaar.length()==2) {
									try {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
									}catch(NumberFormatException e) {
										
									}
								}
								result[43] = bouwjaar;
								
								bouwjaar = ((String) ListA.get(1)[2]).replaceAll("!","1");
								if(bouwjaar.length()==2) {
									try {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
										if(bouwjaar.length()==3) {
											bouwjaar ="200"+jaar;	
										}
									}else {
										bouwjaar ="19"+jaar;
									}
									}catch(NumberFormatException e) {
										
									}
								}
								result[44] = bouwjaar;
							}

						}
						}catch(Exception e) {
							System.out.println("bj_: "+e.getMessage()+"      "+bestand);
							
						}
					
						keyA = "kleur_" + i+"_a";
						keyB = "kleur_" + i+"_b";
						ListA = verzameling.get(keyA);
						ArrayList<Object[]> ListB = verzameling.get(keyB);
						//6 kleur rijen
						
						if (ListA != null) {
							if(ListA.size()>=7) {
								System.out.println("rijen>6   "+bestand);
							}
							if (ListA.size() == 1) {
								result[45] = (String) ListA.get(0)[2];
								result[46] = "";
								result[47] = "";
							} else if (ListA.size() == 2) {
								result[45] = (String) ListA.get(0)[2];
								result[46] = (String) ListA.get(1)[2];
								result[47] = "";
							} else if (ListA.size() > 2) {
								result[45] = (String) ListA.get(0)[2];
								result[46] = (String) ListA.get(2)[2];
								result[47] = (String) ListA.get(1)[2];
							}
							
						}
						
						if (ListB != null) {
							if(ListB.size()>=7) {
								System.out.println("rijen>6   "+bestand);
							}
							if (ListB.size() == 1) {
								result[48] = (String) ListB.get(0)[2];
								result[49] = "";
								result[50] = "";
							} else if (ListB.size() == 2) {
								result[48] = (String) ListB.get(0)[2];
								result[49] = (String) ListB.get(1)[2];
								result[50] = "";
							} else if (ListB.size() > 2) {
								result[48] = (String) ListB.get(0)[2];
								result[49] = (String) ListB.get(2)[2];
								result[50] = (String) ListB.get(1)[2];
							}
							
						}

						keyA = "afstand_" + i;

						ListA = verzameling.get(keyA);

						if (ListA != null) {
							
							if (ListA.size() == 1) {
								result[51] = (String) ListA.get(0)[2];
								result[52] = "";
								 
							} else if (ListA.size() == 2 || ListA.size() == 4) {
								result[51] = (String) ListA.get(0)[2];
								result[52] = (String) ListA.get(1)[2];
								
							}
							
						}
						
						row = sheet.getRow(ArmRow);

						if (row == null) {
							row = sheet.createRow(ArmRow);
						}

						cell = row.getCell(0);
						if (cell == null) {
							cell = row.createCell(0);
						}
						cell.setCellValue(results[1]); // bestandsnaam
						for (int x = 0; x < result.length; x++) {
							if (result[x] != null && !result[x].isBlank()) {
								cell = row.getCell(x + 1);
								if (cell == null) {
									cell = row.createCell(x + 1);
								}
								cell.setCellValue(result[x].strip());
								// if(x==8 || x==11) {
								// cell.setCellStyle(cs); // word wrap
								// }
							}
						}

						ArmRow++;
					}
				}
			}
			}catch(Exception e) {
				System.out.println("catch tab 2: "+e.getMessage());
			}
			//////////////////////////////////////////////////////////////////////////////
			try {
			ArrayList<Object[]> temp = verzameling.get("Art.code plaatsing");

			sheet = workbook.getSheetAt(3); // Werkzaamheden

			if (temp != null) {
				for (int i = 0; i < temp.size(); i++) {
					if(!((String) temp.get(i)[2]).contentEquals("Art.code plaatsing")&&
							!((String) temp.get(i)[2]).contentEquals("Indx.")&&
							!((String) temp.get(i)[2]).contentEquals("Pl.code")) {
					String[] result2 = new String[4];
					result2[0] = results[7]; // object
					result2[1] = (String) temp.get(i)[2]; // Art.code
					String yPos = String.valueOf(((Double) temp.get(i)[1]));// y pos of value, equal y pos betekent
																			// zelfde rij

					ArrayList<Object[]> temp2 = verzameling.get("aant.");
					if (temp2 != null) {
						for (int z = 0; z < temp2.size(); z++) {
							if ((String.valueOf(temp2.get(z)[1])).equals(yPos)) {
								result2[2] = (String) temp2.get(z)[2];
								break;
							}
						}
					}

					temp2 = verzameling.get("omschrijving werkzaamheden");
					if (temp2 != null) {
						for (int z = 0; z < temp2.size(); z++) {
							// System.out.println(temp2.get(z)[1]+" "+yPos+"
							// "+(String.valueOf(temp2.get(z)[1])).equals(yPos));
							if ((String.valueOf(temp2.get(z)[1])).equals(yPos)) {
								result2[3] = (String) temp2.get(z)[2];

								break;
							}
						}
					}

					row = sheet.getRow(WerkzaamhedenRow);

					if (row == null) {
						row = sheet.createRow(WerkzaamhedenRow);
					}

					cell = row.getCell(0);
					if (cell == null) {
						cell = row.createCell(0);
					}
					cell.setCellValue(results[1]); // bestandsnaam
					for (int x = 0; x < result2.length; x++) {
						if (result2[x] != null && !result2[x].equals("")) {
							cell = row.getCell(x + 1);
							if (cell == null) {
								cell = row.createCell(x + 1);
							}
							cell.setCellValue(result2[x]);
						}
					}

					WerkzaamhedenRow++;
					}
				}
			}
			}catch(Exception e) {
				System.out.println("catch tab 3: "+e.getMessage());
			}
			//////////////////////////////////////////////////////////////////////////////
			try {
			ArrayList<Object[]>  temp = verzameling.get("artikel");

			sheet = workbook.getSheetAt(4); // as-Built
			if (temp != null) {
				for (int i = 0; i < temp.size(); i++) {
					
					String[] result2 = new String[6];
					result2[0] = results[7]; // object
					result2[1] = (String) temp.get(i)[2]; // Artikel
					Double yPos = (Double) temp.get(i)[1];// y pos of value, equal y pos betekent
																			// zelfde rij

					ArrayList<Object[]> temp2 = verzameling.get("Nw");
					if (temp2 != null) {
						for (int z = 0; z < temp2.size(); z++) {
							if ((Double)temp2.get(z)[1]>yPos-0.1 &&(Double)temp2.get(z)[1]<yPos+0.1) {
								result2[2] = (String) temp2.get(z)[2];
								break;
							}
						}
					}

					temp2 = verzameling.get("Tot");
					if (temp2 != null) {
						for (int z = 0; z < temp2.size(); z++) {
							if ((Double)temp2.get(z)[1]>yPos-0.1 &&(Double)temp2.get(z)[1]<yPos+0.1) {
								result2[3] = (String) temp2.get(z)[2];
								break;
							}
						}
					}

					temp2 = verzameling.get("Omschrijving");
					if (temp2 != null) {
						
						for (int z = 0; z < temp2.size(); z++) {
							if ((Double)temp2.get(z)[1]>yPos-0.1 &&(Double)temp2.get(z)[1]<yPos+0.1) {
								
								if(result2[4]==null||result2[4].strip().contentEquals("")) {
									result2[4] = ((String) temp2.get(z)[2]).strip();
								}else {
									
									result2[4] =result2[4].strip()+" "+ ((String) temp2.get(z)[2]).strip();
								
								}
								
							}
						}
					}
					try {
					temp2 = verzameling.get("Bouwjr");
					if (temp2 != null) {
						for (int z = 0; z < temp2.size(); z++) {
							if ((Double)temp2.get(z)[1]>yPos-0.1 &&(Double)temp2.get(z)[1]<yPos+0.1) {
								String bouwjaar = (String) temp2.get(z)[2];
								if(bouwjaar.length()==2) {
									if(bouwjaar.equals("bj")) {
										bouwjaar ="";
									}else {
									try {
									int jaar = Integer.valueOf(bouwjaar);
									if(jaar>=0 && jaar <=21) {
										bouwjaar ="20"+jaar;
									}else {
										bouwjaar ="19"+jaar;
									}
									}catch(NumberFormatException e) {
										System.out.println("bouwjaar: "+bouwjaar);
									}
									}
								}
								result2[5] = bouwjaar;
								break;
							}
						}
					}
					}catch(Exception e) {
						System.out.println("Bouwjr: "+e.getMessage());
					}
					row = sheet.getRow(AsBuiltRow);

					if (row == null) {
						row = sheet.createRow(AsBuiltRow);
					}

					cell = row.getCell(0);
					if (cell == null) {
						cell = row.createCell(0);
					}
					cell.setCellValue(results[1]); // bestandsnaam
					for (int x = 0; x < result2.length; x++) {
						if (result2[x] != null && !result2[x].equals("")) {
							cell = row.getCell(x + 1);
							if (cell == null) {
								cell = row.createCell(x + 1);
							}
							cell.setCellValue(result2[x]);
						}
					}

					AsBuiltRow++;
				}
			}
			}catch(Exception e) {
				System.out.println("catch tab 4: "+e.getMessage());
			}

		} else {
			System.out.println("Doel excel bestaat niet");
		}
		}catch(Exception e) {
			System.out.println("Printen: "+e.getMessage());
		}

	}



	@SuppressWarnings({ "rawtypes", "unchecked" })
	private static void sorteer() {
		Iterator it = verzameling.entrySet().iterator();

		while (it.hasNext()) {
			HashMap.Entry<String, ArrayList<Object[]>> pair = (HashMap.Entry) it.next();
			ArrayList<Object[]> tempList = pair.getValue();
			ArrayList<String> tempList2 = new ArrayList<String>();
			// temp[0] = x;
			// temp[1] = y;
			// temp[2] = value;
			//String[] woordResult = new String[1];
			if (pair.getKey().startsWith("bord")) {
				if (tempList.size() > 1) {
					// sorteer alleen groepen met meer dan 1 entry
					boolean done = false;
					while (done == false) {
						done = true;
						for (int i = 1; i < tempList.size(); i++) {

							// sorteer y van hoog naar laag
							double temp1 = (double) tempList.get(i - 1)[1];
							double temp2 = (double) tempList.get(i)[1];

							double temp3 = (double) tempList.get(i - 1)[0];
							double temp4 = (double) tempList.get(i)[0];
							if (temp2 > temp1) {
								// swap
								done = false;
								Collections.swap(tempList, i, i - 1);
							} else if (temp2 == temp1 && temp3 > temp4) {
								// swap
								done = false;
								Collections.swap(tempList, i, i - 1);
							}
						}
					}
					String woord = "";
					double y = 0;
					for (int i = 0; i < tempList.size(); i++) {
						if (i == 0) {
							y = (double) tempList.get(i)[1];

							woord = woord + tempList.get(i)[2];
						} else if (y == (double) tempList.get(i)[1]) {

							if ((isNumeric((String) tempList.get(i)[2])
									&& !isNumeric(woord.substring(woord.length() - 1)))) {
								woord = woord + " " + tempList.get(i)[2];
							}else if(i!=0 && Character.compare(woord.charAt(woord.length()-1),'-') !=0 && Character.isUpperCase(((String)tempList.get(i)[2]).charAt(0))){
								woord = woord + " " + tempList.get(i)[2];
								
							}else {
								
								woord = woord + tempList.get(i)[2];
								
							}
						} else {
							// System.out.println(woord); //compleet woord	
							if(woord.contentEquals("()")==false && woord.strip().equals("")==false) {
							tempList2.add(woord);
							}
							y = (double) tempList.get(i)[1];
							woord = (String) tempList.get(i)[2];
							

						}

					}
					if(woord.contentEquals("()")==false && woord.strip().equals("")==false) {
					tempList2.add(woord);
					}
					// System.out.println(woord); //compleet woord
				} else if (tempList.size() == 1) {
					// System.out.println((String) tempList.get(0)[2]); //compleet woord
					// woordResult[0] = (String) tempList.get(0)[2];
					if(((String) tempList.get(0)[2]).strip().equals(results[6])||
							((String) tempList.get(0)[2]).contentEquals("()")) {
						
					}else {
					tempList2.add((String) tempList.get(0)[2]);
					}
				}
			} else {

				if (tempList.size() > 1) {
					// sorteer alleen groepen met meer dan 1 entry
					boolean done = false;
					while (done == false) {
						done = true;
						for (int i = 1; i < tempList.size(); i++) {

							// sorteer y van hoog naar laag
							double temp1 = (double) tempList.get(i - 1)[1];
							double temp2 = (double) tempList.get(i)[1];
							if (temp2 > temp1) {
								// swap
								done = false;
								Collections.swap(tempList, i, i - 1);
							}else if(temp2==temp1) {
								temp1 = (double) tempList.get(i - 1)[0];
								temp2 = (double) tempList.get(i)[0];
								if (temp1 > temp2) {
									// swap
									done = false;
									Collections.swap(tempList, i, i - 1);
								}
							}
						}
					}

					
				}
			}
			// verzameling.put(pair.getKey(), tempList);
			if (tempList2.size() > 0) {
				verzameling2.put("woord_" + pair.getKey(), tempList2);
			}
		}

	}

	public static boolean isNumeric(String strNum) {
		if (strNum == null) {
			return false;
		}
		try {
			int i = Integer.parseInt(strNum);
		} catch (NumberFormatException nfe) {
			return false;
		}
		return true;
	}
	
	public static boolean isNumeric2(String strNum) {
		if (strNum == null) {
			return false;
		}
		try {
			if(strNum.equals(",")) {
				return true;
			}else {
			double i = Double.parseDouble(strNum.replaceAll(",", "."));
			}
		} catch (NumberFormatException nfe) {
			return false;
		}
		return true;
	}

	private static void voegToe(double x, double y, String value, String key) {
		if(value.strip().equals("")) {
			
		}else {
		ArrayList<Object[]> tempList = verzameling.get(key);
		if (tempList == null) {
			tempList = new ArrayList<Object[]>();
		}
		Object[] temp = new Object[3];
		temp[0] = x;
		temp[1] = y;
		temp[2] = value;
		tempList.add(temp);

		verzameling.put(key, tempList);
		}
	}

	private static void voegToe(double x, double y, String value, String key, double X1, double X2, double Y1, double Y2) {
		//X1 = x link
		//X2 = x rechts
		//Y1 = y laag
		//Y2 = y hoog
		if(value.strip().equals("")) {
			
		}else {
		double xtot = X2-X1;
		double ytot = Y2-Y1;
		String loc ="";
		if(!value.contains("(links)")&&!value.contains("(rechts)")){
			
			if(x>=X1 && x<=X1+(xtot/3)) {
				loc = "links ";
			}else if(x>=X1+(xtot/3) && x<=X1+(xtot/3)+(xtot/3)) {
				loc = "midden ";
			}else if(x>=X1+(xtot/3)+(xtot/3) && x<=X1+(xtot/3)+(xtot/3)+(xtot/3)) {
				loc = "rechts ";
			}
			
			if(y>=Y1 && y<=Y1+(ytot/3)) {
				loc = loc+"onder";
			}else if(y>=Y1+(ytot/3) && y<=Y1+(ytot/3)+(ytot/3)) {
				loc = loc+"midden";
			}else if(y>=Y1+(ytot/3)+(ytot/3) && y<=Y1+(ytot/3)+(ytot/3)+(ytot/3)) {
				loc = loc+"boven";
			}
			
			if(loc.equals("midden midden")) {
				loc = "midden";
			}else if(loc.equals("midden boven")) {
				loc = "boven";
			}else if(loc.equals("midden onder")) {
				loc = "onder";
			}else if(loc.equals("links midden")) {
				loc = "links";
			}else if(loc.equals("rechts midden")) {
				loc = "rechts";
			}
		}
		ArrayList<Object[]> tempList = verzameling.get(key);
		if (tempList == null) {
			tempList = new ArrayList<Object[]>();
		}
		Object[] temp = new Object[3];
		temp[0] = x;
		temp[1] = y;
		if(loc.equals("")) {
			temp[2] = value;	
		}else {
		temp[2] = value+" ("+loc+")";
		}
		tempList.add(temp);

		verzameling.put(key, tempList);
		}
		}
	
	@SuppressWarnings("resource")
	private static void readExtracts() throws FileNotFoundException {
		int counter2=1;
		File[] fileList = inputBestandenVoorRun;
		if (fileList == null) {
			File extractFiles = new File("Bron");
			fileList = extractFiles.listFiles();
		}
		if (fileList == null || fileList.length == 0) {
			System.out.println("Geen extractbestanden gevonden");
			return;
		}
		//for (int i = fileList.length-1; i> 900; i--) {
		//for (int i = 200; i> 150; i--) {
		
		for (int i = 0; i < fileList.length; i++) { //TODO //nu op 400
		//for (int i = 900; i < fileList.length; i++) {
			if (fileList[i] == null || fileList[i].isFile() == false) {
				continue;
			}
			objecten.clear();
			verzameling.clear();
			verzameling2.clear();
			klokstanden.clear();
			System.out.println(fileList[i].getName()+"     nr: " +counter2);
			counter2++;
			FileInputStream fsIP = null;
			BufferedInputStream bIP = null;
			Workbook workbook = null;
			Sheet sheet;
			Row row;
			Cell cell;

			int[] Indexes = new int[38];

			try {
				fsIP = new FileInputStream(fileList[i]);
				bIP = new BufferedInputStream(fsIP);
				ZipSecureFile.setMinInflateRatio(0.008);
				if (fileList[i].getName().toLowerCase().endsWith(".xlsx")) {
					workbook = new XSSFWorkbook(bIP);
				} else {
					workbook = new HSSFWorkbook(bIP);
				}

				sheet = workbook.getSheetAt(0);

				row = sheet.getRow(0);
				int count = 0;
				cell = row.getCell(count);
				Indexes[26] = -1;
				while (cell != null) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().strip().equals("File Name")) {
							Indexes[0] = count;
						} else if (cell.getStringCellValue().strip().equals("File Location")) {
							Indexes[1] = count;
						} else if (cell.getStringCellValue().strip().equals("Name")) {
							Indexes[2] = count;
						} else if (cell.getStringCellValue().strip().equals("Position X")) {
							Indexes[3] = count;
						} else if (cell.getStringCellValue().strip().equals("Position Y")) {
							Indexes[4] = count;
						} else if (cell.getStringCellValue().strip().equals("Position Z")) {
							Indexes[5] = count;
						} else if (cell.getStringCellValue().strip().equals("Start X")) {
							Indexes[6] = count;
						} else if (cell.getStringCellValue().strip().equals("Start Y")) {
							Indexes[7] = count;
						} else if (cell.getStringCellValue().strip().equals("Start Z")) {
							Indexes[8] = count;
						} else if (cell.getStringCellValue().strip().equals("End X")) {
							Indexes[9] = count;
						} else if (cell.getStringCellValue().strip().equals("End Y")) {
							Indexes[10] = count;
						} else if (cell.getStringCellValue().strip().equals("End Z")) {
							Indexes[11] = count;
						} else if (cell.getStringCellValue().strip().equals("DATUM")) {
							Indexes[12] = count;
						} else if (cell.getStringCellValue().strip().equals("PLAATDEEL")) {
							Indexes[13] = count;
						} else if (cell.getStringCellValue().strip().equals("GETEKEND")) {
							Indexes[14] = count;
						} else if (cell.getStringCellValue().strip().equals("GEZIEN")) {
							Indexes[15] = count;
						} else if (cell.getStringCellValue().strip().equals("REVISIE")) {
							Indexes[16] = count;
						} else if (cell.getStringCellValue().strip().equals("SOORT")) {
							Indexes[17] = count;
						} else if (cell.getStringCellValue().strip().equals("SOORT_HANDW")) {
							Indexes[18] = count;
						} else if (cell.getStringCellValue().strip().equals("Value")) {
							Indexes[19] = count;
						} else if (cell.getStringCellValue().strip().equals("Layer")) {
							Indexes[20] = count;
						} else if (cell.getStringCellValue().strip().equals("Length")) {
							Indexes[21] = count;
						} else if (cell.getStringCellValue().strip().equals("Rotation")) {
							Indexes[22] = count;
						} else if (cell.getStringCellValue().strip().equals("File Modified")) {
							Indexes[23] = -1;
						} else if (cell.getStringCellValue().strip().equals("CAMTAG")) {
							Indexes[24] = count;
						}else if (cell.getStringCellValue().strip().equals("Closed")) {
							Indexes[25] = -1;
						}else if (cell.getStringCellValue().strip().equals("_\\U+002E\\U+002E\\U+002E")) {//...
							
							Indexes[26] = count;
						}else if (cell.getStringCellValue().strip().equals("Prompt")) {
							Indexes[27] = count;
						}else if (cell.getStringCellValue().strip().equals("Tag")) {
							Indexes[28] = count;
						}else if(cell.getStringCellValue().strip().equals("CADFIL")) {
							Indexes[29] = count;
						}else if(cell.getStringCellValue().strip().equals("EXTRAINFO")) {
							Indexes[30] = count;
						}else if(cell.getStringCellValue().strip().equals("EXTRAINFO(1)")) {
							Indexes[31] = count;
						}else if(cell.getStringCellValue().strip().equals("UPPERCASE")) {
							Indexes[32] = count;
						}else if(cell.getStringCellValue().strip().equals("LOWERCASE")) {
							Indexes[33] = count;
						}else if(cell.getStringCellValue().strip().equals("LTYPE")) {
							Indexes[34] = count;
						}else if(cell.getStringCellValue().strip().equals("BW")) {
							Indexes[35] = count;
						}else if(cell.getStringCellValue().strip().equals("BH")) {
							Indexes[36] = count;
						}else if(cell.getStringCellValue().strip().equals("Contents")) {
							Indexes[37] = count;
						}

					}

					count++;
					cell = row.getCell(count);
				}

				int counter = 1;
				row = sheet.getRow(counter);
				while (row != null) {

					String[] temp = new String[Indexes.length];
					cell = row.getCell(0);
					if (cell == null || cell.getCellType() == CellType.BLANK || (cell.getCellType() == CellType.STRING
							&& cell.getStringCellValue().strip().equals(""))) {
						break;
					}
					for (int z = 0; z < temp.length; z++) {
						if(Indexes[z] != -1) {
							cell = row.getCell(Indexes[z]);
							if (cell != null) {
								if (cell.getCellType() == CellType.STRING) {
									temp[z] = cell.getStringCellValue();
									
								} else if (cell.getCellType() == CellType.NUMERIC) {
									temp[z] = String.valueOf((int) cell.getNumericCellValue());
								} else if (cell.getCellType() == CellType.BLANK) {
									temp[z] = "";
								} else if (cell.getCellType() == CellType._NONE) {
									temp[z] = "";
								}
	
							} else {
								temp[z] = "";
							}
						}else {
							temp[z] = "-1";				
						}
					}

					String Key = temp[1] + "\\" + temp[0]; // pad+bestandsnaam
					ArrayList<String[]> tempList = objecten.get(Key);
					if (tempList == null) {
						tempList = new ArrayList<String[]>();
					}
					tempList.add(temp);
					objecten.put(Key, tempList);

					counter++;
					row = sheet.getRow(counter);
				}

				workbook.close();
				bIP.close();
				fsIP.close();
				workbook = null;
				sheet = null;
				bIP = null;
				fsIP = null;

				try {
				verwerkExtracts();
				}catch(Exception e) {
					System.out.println("test 4: "+e.getMessage());
				}
				
				System.gc();
				
			} catch (Exception e) {
				throw new RuntimeException("Fout bij verwerken extract: " + fileList[i].getName(), e);
			} finally {
				closeQuietly(workbook);
				closeQuietly(bIP);
				closeQuietly(fsIP);
			}
			System.gc();
		}

	}

}
