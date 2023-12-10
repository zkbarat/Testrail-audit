package org.example;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Scanner;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class TestRailReportGeneratorTest {

  public static void main(String[] args) {
    Scanner scanner = new Scanner(System.in);
    System.out.print("Enter the start date (MM/dd/yyyy): ");
    String startDateString = scanner.nextLine();
    long startTime = parseUnixTimestamp(startDateString);
    System.out.print("Enter the end date (MM/dd/yyyy): ");
    String endDateString = scanner.nextLine();
    long endTime = parseUnixTimestamp(endDateString);
    System.out.print("Enter the base URL:");
    String baseUrl = scanner.nextLine();
    System.out.print("Enter your email:");
    String email = scanner.nextLine();
    System.out.print("Enter your API key:");
    String apiKey = scanner.nextLine();
    String absPath = "/Users/zkbarat/IdeaProjects/untitled2/src/main/java/org/example/userMappings.Json";
    Map<String, Integer> userMappings = loadUserMappingsFromFile(absPath);
    Map<Integer, String> userMappingsIntegerToString = new HashMap<>();
    for (Map.Entry<String, Integer> entry : userMappings.entrySet()) {
      userMappingsIntegerToString.put(entry.getValue(), entry.getKey());
    }
    System.out.print("Enter QSA name:");
    String testedByUserInput = scanner.nextLine();
    List<String> testedByUsernames = Arrays.asList(testedByUserInput.split(","));
    List<Integer> preferredProjectIds = new ArrayList<>();
    System.out.print(
        "Enter preferred project IDs (comma-separated) or 'no' to fetch all projects: ");
    String preferredProjectIdsInput = scanner.nextLine();
    List<Integer> preferredTestRunIds = new ArrayList<>();
    System.out.print(
        "Enter preferred test run IDs (comma-separated) or 'no' to fetch all test runs: ");
    String preferredTestRunsInput = scanner.nextLine();
    System.out.println("Please wait while we generate your report......");
    if (!preferredProjectIdsInput.equalsIgnoreCase("no")) {
      String[] preferredIdsArray = preferredProjectIdsInput.split(",");
      for (String id : preferredIdsArray) {
        preferredProjectIds.add(Integer.parseInt(id.trim()));
      }
    }
    if (!preferredTestRunsInput.equalsIgnoreCase("no")) {
      String[] preferredIdsArray = preferredTestRunsInput.split(",");
      for (String id : preferredIdsArray) {
        preferredTestRunIds.add(Integer.parseInt(id.trim()));
      }
    }
    HttpClient httpClient = HttpClient.newHttpClient();
    try {
      // Fetch projects
      JSONArray projects;
      if (!preferredProjectIds.isEmpty()) {
        projects = new JSONArray();
        for (int projectId : preferredProjectIds) {
          JSONObject projectObj = new JSONObject();
          projectObj.put("id", projectId);
          projects.put(projectObj);
        }
      } else {
        projects = fetchProjects(httpClient, baseUrl, email, apiKey);
      }
      // Fetch data from all projects and test runs
      List<JSONObject> selectedTestCases = Collections.synchronizedList(new ArrayList<>());
      for (int i = 0; i < projects.length(); i++) {
        JSONObject project = projects.getJSONObject(i);
        int projectId = project.getInt("id");
        JSONArray allTestRuns;
        if (!preferredTestRunIds.isEmpty()) {
          allTestRuns = new JSONArray();
          for (int testRunId : preferredTestRunIds) {
            JSONObject testRunObj = new JSONObject();
            testRunObj.put("id", testRunId);
            allTestRuns.put(testRunObj);
          }
        } else {
          allTestRuns = fetchAllTestRuns(httpClient, baseUrl, email, apiKey, projectId);
        }
        System.out.println("Testruns fetched, please wait while we fetch results from testruns");
        ExecutorService executorService = Executors.newFixedThreadPool(
            Runtime.getRuntime().availableProcessors());
        for (int j = 0; j < allTestRuns.length(); j++) {
          JSONObject testRun = allTestRuns.getJSONObject(j);
          int testRunId = testRun.getInt("id");
          for (String testedByUsername : testedByUsernames) {
            List<Integer> mappedUserIds = new ArrayList<>();
            int userId = userMappings.getOrDefault(testedByUsername, -1);
            if (userId != -1) {
              mappedUserIds.add(userId);
              executorService.execute(() -> {
                JSONArray results = fetchResultsForTestRun(httpClient, baseUrl, email, apiKey,
                    testRunId, mappedUserIds);
                List<JSONObject> selectedTestCasesForRun = selectRandomTestCases(results, 0.1,
                    startTime, endTime);
                synchronized (selectedTestCases) {
                  selectedTestCases.addAll(selectedTestCasesForRun);
                }
              });
            }
          }
        }
        executorService.shutdown();
        try {
          executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
        } catch (InterruptedException e) {
          e.printStackTrace();
        }
      }
      System.out.println("Test results fetched, please wait while we generating excel");
      // Generate Excel report for selected test cases
      ExecutorService reportExecutorService = Executors.newFixedThreadPool(
          Runtime.getRuntime().availableProcessors());
      reportExecutorService.execute(() ->
          generateExcelReport(selectedTestCases, userMappingsIntegerToString, baseUrl, email,
              apiKey));
      reportExecutorService.shutdown();
      try {
        reportExecutorService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
      } catch (InterruptedException e) {
        e.printStackTrace();
      }
      System.out.println("Excel report is generated successfully.");
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      scanner.close();
    }
  }

  private static long parseUnixTimestamp(String dateString) {
    LocalDate localDate = LocalDate.parse(dateString, DateTimeFormatter.ofPattern("MM/dd/yyyy"));
    return localDate.atStartOfDay(ZoneOffset.UTC).toEpochSecond();
  }

  private static Map<String, Integer> loadUserMappingsFromFile(String filePath) {
    Map<String, Integer> userMappings = new HashMap<>();
    try (Reader reader = new FileReader(filePath)) {
      Gson gson = new Gson();
      Map<String, Map<String, Integer>> data = gson.fromJson(reader,
          new TypeToken<Map<String, Map<String, Integer>>>() {
          }.getType());
      userMappings = data.get("userMappings");
    } catch (IOException e) {
      e.printStackTrace();
    }
    return userMappings;
  }

  private static Map<Integer, String> fetchStatuses(HttpClient httpClient, String baseUrl,
      String email, String apiKey) {
    Map<Integer, String> statusMappings = new HashMap<>();
    String apiUrl = baseUrl + "/index.php?/api/v2/get_statuses";
    HttpRequest request = HttpRequest.newBuilder()
        .uri(URI.create(apiUrl))
        .header("Content-Type", "application/json")
        .header("Authorization", "Basic " + java.util.Base64.getEncoder()
            .encodeToString((email + ":" + apiKey).getBytes()))
        .build();
    try {
      HttpResponse<String> response = httpClient.send(request,
          HttpResponse.BodyHandlers.ofString());
      JSONArray statusArray = new JSONArray(response.body());
      for (int i = 0; i < statusArray.length(); i++) {
        JSONObject statusObject = statusArray.getJSONObject(i);
        int statusId = statusObject.getInt("id");
        String statusName = statusObject.getString("label");
        statusMappings.put(statusId, statusName);
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return statusMappings;
  }

  private static JSONArray fetchProjects(HttpClient httpClient, String baseUrl, String email,
      String apiKey) {
    String apiUrl = baseUrl + "/index.php?/api/v2/get_projects/";
    HttpRequest request = HttpRequest.newBuilder()
        .uri(URI.create(apiUrl))
        .header("Content-Type", "application/json")
        .header("Authorization", "Basic " + java.util.Base64.getEncoder()
            .encodeToString((email + ":" + apiKey).getBytes()))
        .build();
    try {
      HttpResponse<String> response = httpClient.send(request,
          HttpResponse.BodyHandlers.ofString());
      return new JSONArray(response.body());
    } catch (Exception e) {
      e.printStackTrace();
      return new JSONArray();
    }
  }

  private static JSONArray fetchAllTestRuns(HttpClient httpClient, String baseUrl, String email,
      String apiKey, int projectId) {
    JSONArray allTestRuns = new JSONArray();
    int pageSize = 250;
    int page = 1;
    LocalDateTime currentDate = LocalDateTime.now();
    LocalDateTime sixMonth = currentDate.minusMonths(6);
    while (true) {
      // Fetch standalone test runs
      String standaloneApiUrl =
          baseUrl + "/index.php?/api/v2/get_runs/" + projectId + "&limit=" + pageSize + "&offset="
              + (pageSize * (page - 1));
      JSONArray standaloneTestRuns = fetchTestRuns(httpClient, standaloneApiUrl, email, apiKey,
          sixMonth, currentDate);
      allTestRuns.putAll(standaloneTestRuns);
      // Fetch test runs within test plans
      String planApiUrl =
          baseUrl + "/index.php?/api/v2/get_plans/" + projectId + "&limit=" + pageSize + "&offset="
              + (pageSize * (page - 1));
      JSONArray testPlans = fetchTestPlans(httpClient, planApiUrl, email, apiKey);
      JSONArray planTestRuns = fetchTestRunsFromPlans(httpClient, baseUrl, email, apiKey, testPlans,
          sixMonth, currentDate);
      allTestRuns.putAll(planTestRuns);
      if (standaloneTestRuns.isEmpty() && planTestRuns.isEmpty()) {
        break;
      }
      page++;
    }
    return allTestRuns;
  }

  private static JSONArray fetchTestRuns(HttpClient httpClient, String apiUrl, String email,
      String apiKey, LocalDateTime sixMonth, LocalDateTime currentDate) {
    JSONArray testRuns = new JSONArray();
    try {
      HttpRequest request = HttpRequest.newBuilder()
          .uri(URI.create(apiUrl))
          .header("Content-Type", "application/json")
          .header("Authorization", "Basic " + java.util.Base64.getEncoder()
              .encodeToString((email + ":" + apiKey).getBytes()))
          .build();
      HttpResponse<String> response = httpClient.send(request,
          HttpResponse.BodyHandlers.ofString());
      JSONArray currentPageTestRuns = new JSONArray(response.body());
      for (int i = 0; i < currentPageTestRuns.length(); i++) {
        JSONObject testRun = currentPageTestRuns.getJSONObject(i);
        long createdTimestamp = testRun.getLong("created_on");
        LocalDateTime createdDateTime = Instant.ofEpochSecond(createdTimestamp)
            .atZone(ZoneId.systemDefault()).toLocalDateTime();
        if (!createdDateTime.isBefore(sixMonth) && !createdDateTime.isAfter(currentDate)) {
          testRuns.put(testRun);
        }
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return testRuns;
  }

  private static JSONArray fetchTestPlans(HttpClient httpClient, String apiUrl, String email,
      String apiKey) {
    JSONArray testPlans = new JSONArray();
    try {
      HttpRequest request = HttpRequest.newBuilder()
          .uri(URI.create(apiUrl))
          .header("Content-Type", "application/json")
          .header("Authorization", "Basic " + java.util.Base64.getEncoder()
              .encodeToString((email + ":" + apiKey).getBytes()))
          .build();
      HttpResponse<String> response = httpClient.send(request,
          HttpResponse.BodyHandlers.ofString());
      testPlans = new JSONArray(response.body());
    } catch (Exception e) {
      e.printStackTrace();
    }
    return testPlans;
  }

  private static JSONArray fetchTestRunsFromPlans(HttpClient httpClient, String baseUrl,
      String email, String apiKey, JSONArray testPlans, LocalDateTime sixMonth,
      LocalDateTime currentDate) {
    JSONArray planTestRuns = new JSONArray();
    for (int i = 0; i < testPlans.length(); i++) {
      JSONObject plan = testPlans.getJSONObject(i);
      int planId = plan.getInt("id");
      String planApiUrl = baseUrl + "/index.php?/api/v2/get_plan/" + planId;
      try {
        HttpRequest request = HttpRequest.newBuilder()
            .uri(URI.create(planApiUrl))
            .header("Content-Type", "application/json")
            .header("Authorization", "Basic " + java.util.Base64.getEncoder()
                .encodeToString((email + ":" + apiKey).getBytes()))
            .build();
        HttpResponse<String> response = httpClient.send(request,
            HttpResponse.BodyHandlers.ofString());
        JSONObject planDetails = new JSONObject(response.body());
        JSONArray entries = planDetails.getJSONArray("entries");
        for (int j = 0; j < entries.length(); j++) {
          JSONObject entry = entries.getJSONObject(j);
          JSONArray runs = entry.getJSONArray("runs");
          planTestRuns.putAll(filterTestRunsByDate(runs, sixMonth, currentDate));
        }
      } catch (Exception e) {
        e.printStackTrace();
      }
    }
    return planTestRuns;
  }

  private static JSONArray filterTestRunsByDate(JSONArray runs, LocalDateTime sixMonth,
      LocalDateTime currentDate) {
    JSONArray filteredRuns = new JSONArray();
    for (int i = 0; i < runs.length(); i++) {
      JSONObject run = runs.getJSONObject(i);
      long createdTimestamp = run.getLong("created_on");
      LocalDateTime createdDateTime = Instant.ofEpochSecond(createdTimestamp)
          .atZone(ZoneId.systemDefault()).toLocalDateTime();
      if (!createdDateTime.isBefore(sixMonth) && !createdDateTime.isAfter(currentDate)) {
        filteredRuns.put(run);
      }
    }
    return filteredRuns;
  }

  private static JSONArray fetchResultsForTestRun(HttpClient httpClient, String baseUrl,
      String email, String apiKey, int testRunId, List<Integer> testedByUserIds) {
    String apiUrl = baseUrl + "/index.php?/api/v2/get_results_for_run/" + testRunId;
    JSONArray allResults = new JSONArray();
    try {
      int offset = 0;
      int limit = 250;
      while (true) {
        String paginatedApiUrl = apiUrl + "&created_by=" + String.join(",",
            testedByUserIds.stream().map(Object::toString).collect(Collectors.toList())) +
            "&limit=" + limit + "&offset=" + offset;
        HttpRequest request = HttpRequest.newBuilder()
            .uri(URI.create(paginatedApiUrl))
            .header("Content-Type", "application/json")
            .header("Authorization", "Basic " + java.util.Base64.getEncoder()
                .encodeToString((email + ":" + apiKey).getBytes()))
            .build();
        HttpResponse<String> response = httpClient.send(request,
            HttpResponse.BodyHandlers.ofString());
        JSONArray jsonResponse = new JSONArray(response.body());
        for (int i = 0; i < jsonResponse.length(); i++) {
          JSONObject result = jsonResponse.getJSONObject(i);
          allResults.put(result);
        }
        if (jsonResponse.length() < limit) {
          // Break the loop if there are fewer results than the limit, indicating the end of the data
          break;
        }
        offset += limit;
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return allResults;
  }

  private static List<JSONObject> selectRandomTestCases(JSONArray testCases, double percentage,
      long startTime, long endTime) {
    List<JSONObject> selectedTestCases = new ArrayList<>();
    Random random = new Random();
    for (int i = 0; i < testCases.length(); i++) {
      JSONObject testCase = testCases.getJSONObject(i);
      long createdTimestamp = testCase.getLong("created_on");
      LocalDate createdDate = Instant.ofEpochSecond(createdTimestamp).atZone(ZoneId.systemDefault())
          .toLocalDate();
      if (!createdDate.isBefore(
          Instant.ofEpochSecond(startTime).atZone(ZoneId.systemDefault()).toLocalDate())
          && !createdDate.isAfter(
          Instant.ofEpochSecond(endTime).atZone(ZoneId.systemDefault()).toLocalDate())) {
        selectedTestCases.add(testCase);
      }
    }
    int selectedCount = (int) Math.min(Math.ceil(selectedTestCases.size() * percentage),
        selectedTestCases.size());
    while (selectedTestCases.size() > selectedCount) {
      int randomIndex = random.nextInt(selectedTestCases.size());
      selectedTestCases.remove(randomIndex);
    }
    return selectedTestCases;
  }

  private static void generateExcelReport(List<JSONObject> testCases,
      Map<Integer, String> userMappings, String baseUrl, String email, String apiKey) {
    try (Workbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("TestRail Report");
      // Create header row
      Row headerRow = sheet.createRow(0);
      headerRow.createCell(0).setCellValue("TestcaseID");
      headerRow.createCell(1).setCellValue("Title");
      headerRow.createCell(2).setCellValue("Status");
      headerRow.createCell(3).setCellValue("Tested By");
      headerRow.createCell(4).setCellValue("Defects");
      headerRow.createCell(5).setCellValue("Comments");
      CreationHelper creationHelper = workbook.getCreationHelper();
      CellStyle hyperlinkStyle = workbook.createCellStyle();
      Font hyperlinkFont = workbook.createFont();
      hyperlinkFont.setUnderline(Font.U_SINGLE);
      hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
      hyperlinkStyle.setFont(hyperlinkFont);
      HttpClient httpClient = HttpClient.newHttpClient();
      Map<Integer, String> statusMappings = fetchStatuses(httpClient, baseUrl, email, apiKey);
      // Create the vertical alignment and wrap text cell style
      CellStyle cellStyle = workbook.createCellStyle();
      cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
      cellStyle.setWrapText(true);

      // Populate data rows
      for (int i = 0; i < testCases.size(); i++) {
        JSONObject testCase = testCases.get(i);
        int id = testCase.getInt("test_id");
        JSONObject testCaseDetails = fetchTestCaseDetails(httpClient, baseUrl, email, apiKey, id);
        String title = testCaseDetails.getString("title");
        int statusId = testCaseDetails.getInt("status_id");
        String statusName = statusMappings.getOrDefault(statusId, "Unknown Status");
        int testedBy = testCase.getInt("created_by");
        String testedByUsername = userMappings.getOrDefault(testedBy, "Unknown");
        Object defectValue = testCase.get("defects");
        String defect = "";
        if (defectValue instanceof String) {
          defect = (String) defectValue;
        }
        Object commentValue = testCase.get("comment");
        String comment = "";
        if (commentValue instanceof JSONArray) {
          JSONArray commentArray = (JSONArray) commentValue;
          if (commentArray.length() > 0) {
            JSONObject commentObject = commentArray.getJSONObject(0);
            comment = commentObject.getString("comment");
          }
        } else if (commentValue instanceof String) {
          comment = (String) commentValue;
        }
        Row dataRow = sheet.createRow(i + 1);
        Cell idCell = dataRow.createCell(0);
        idCell.setCellValue("T" + id);
        idCell.setCellStyle(hyperlinkStyle);
        Hyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(baseUrl + "/index.php?/tests/view/" + id);
        idCell.setHyperlink(hyperlink);
        dataRow.createCell(1).setCellValue(title);
        dataRow.getCell(1).setCellStyle(cellStyle);
        sheet.autoSizeColumn(1);
        dataRow.createCell(2).setCellValue(statusName);
        dataRow.getCell(2).setCellStyle(cellStyle);
        dataRow.createCell(3).setCellValue(testedByUsername);
        dataRow.getCell(3).setCellStyle(cellStyle);
        Cell defectCell = dataRow.createCell(4);
        defectCell.setCellValue(defect);
        defectCell.setCellStyle(hyperlinkStyle);
        Hyperlink defectHyperLink = creationHelper.createHyperlink(HyperlinkType.URL);
        defectHyperLink.setAddress(defect);
        sheet.autoSizeColumn(4);
        Cell commentCell = dataRow.createCell(5);
        commentCell.setCellValue(comment);
        commentCell.setCellStyle(cellStyle);
        // Auto-adjust row height to fit comment
        dataRow.setHeightInPoints((4 * sheet.getDefaultRowHeightInPoints()));
        // Auto-fit column width for comment
        sheet.autoSizeColumn(5);
      }
      // Save the workbook
      try (FileOutputStream fileOut = new FileOutputStream("TestRailReport.xlsx")) {
        workbook.write(fileOut);
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static JSONObject fetchTestCaseDetails(HttpClient httpClient, String baseUrl,
      String email, String apiKey, int testId) {
    String apiUrl = baseUrl + "/index.php?/api/v2/get_test/" + testId;
    HttpRequest request = HttpRequest.newBuilder()
        .uri(URI.create(apiUrl))
        .header("Content-Type", "application/json")
        .header("Authorization", "Basic " + java.util.Base64.getEncoder()
            .encodeToString((email + ":" + apiKey).getBytes()))
        .build();
    try {
      HttpResponse<String> response = httpClient.send(request,
          HttpResponse.BodyHandlers.ofString());
      return new JSONObject(response.body());
    } catch (Exception e) {
      e.printStackTrace();
      return new JSONObject();
    }
  }
}
