/**
 * @file lsi_exposure_control.cpp
 * @brief LSI半導体チップ露光装置 上位制御プログラム
 *
 * 担当範囲:
 *   - PLCとのEthernet通信（SLMP/MCプロトコル）
 *   - アライメントカメラ画像処理・位置ズレ算出
 *   - サーボモータ位置補正指令
 *   - 露光パラメータ管理
 *   - 検査結果判定・選別
 *   - ログ・統計管理
 */

#include <iostream>
#include <cstring>
#include <cmath>
#include <vector>
#include <queue>
#include <mutex>
#include <thread>
#include <atomic>
#include <chrono>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <functional>
#include <stdexcept>
#include <algorithm>

#ifdef _WIN32
  #include <winsock2.h>
  #include <ws2tcpip.h>
  #pragma comment(lib, "ws2_32.lib")
#else
  #include <sys/socket.h>
  #include <arpa/inet.h>
  #include <unistd.h>
  #define SOCKET int
  #define INVALID_SOCKET (-1)
  #define SOCKET_ERROR   (-1)
  #define closesocket close
#endif

// ============================================================
// 定数定義
// ============================================================
namespace Const {
    // PLC通信
    constexpr char   PLC_IP[]        = "192.168.1.10";
    constexpr uint16_t PLC_PORT      = 5007;          // SLMP UDPポート
    constexpr int    COMM_TIMEOUT_MS = 1000;

    // PLC デバイスアドレス (D レジスタ)
    constexpr uint16_t D_MISALIGN_X  = 0;    // X軸ズレ量
    constexpr uint16_t D_MISALIGN_Y  = 1;    // Y軸ズレ量
    constexpr uint16_t D_MISALIGN_TH = 2;    // θ軸ズレ量
    constexpr uint16_t D_CORR_X      = 3;    // X補正指令
    constexpr uint16_t D_CORR_Y      = 4;    // Y補正指令
    constexpr uint16_t D_CORR_TH     = 5;    // θ補正指令
    constexpr uint16_t D_STEP        = 10;   // ステップカウンタ
    constexpr uint16_t D_NG_COUNT    = 20;   // NG枚数
    constexpr uint16_t D_OK_COUNT    = 21;   // OK枚数
    constexpr uint16_t D_EXP_TIME    = 30;   // 露光時間(ms)

    // PLC ビットデバイス (M コイル)
    constexpr uint16_t M_INIT_DONE   = 0;    // 初期化完了
    constexpr uint16_t M_ALIGN_TRIG  = 10;   // アライメントトリガ
    constexpr uint16_t M_OK_FLAG     = 11;   // OK判定
    constexpr uint16_t M_NG_FLAG     = 12;   // NG判定
    constexpr uint16_t M_ESTOP       = 20;   // 非常停止

    // アライメント許容値
    constexpr double   ALIGN_TOL_XY  = 5.0;  // XY許容誤差 [μm]
    constexpr double   ALIGN_TOL_TH  = 0.01; // θ許容誤差 [deg]
    constexpr double   ALIGN_MAX_XY  = 50.0; // XY最大補正範囲 [μm]
    constexpr int      ALIGN_RETRY   = 3;    // リトライ回数

    // 露光パラメータ
    constexpr double   EXPOSURE_ENERGY = 25.0;  // 露光エネルギー [mJ/cm²]
    constexpr double   WAVELENGTH_NM   = 193.0; // ArF波長 [nm]

    // 検査閾値
    constexpr double   INSPECT_CD_TOL  = 2.0;   // CD誤差許容 [nm]
    constexpr double   INSPECT_DEF_MAX = 0.1;   // 欠陥面積率許容 [%]
}

// ============================================================
// データ構造
// ============================================================

/** アライメント計測結果 */
struct AlignmentResult {
    double dx_um  = 0.0;    ///< X方向ズレ量 [μm]
    double dy_um  = 0.0;    ///< Y方向ズレ量 [μm]
    double dtheta = 0.0;    ///< 回転ズレ量 [deg]
    double residual_xy = 0.0; ///< 補正後残留XY誤差 [μm]
    bool   valid  = false;  ///< 計測有効フラグ
    std::chrono::system_clock::time_point timestamp;
};

/** 検査結果 */
struct InspectionResult {
    enum class Verdict { OK, NG_CD, NG_DEFECT, NG_ALIGN };
    Verdict verdict = Verdict::NG_DEFECT;
    double  cd_error_nm  = 0.0;    ///< CD（線幅）誤差 [nm]
    double  defect_ratio = 0.0;    ///< 欠陥面積率 [%]
    int     defect_count = 0;      ///< 欠陥個数
    bool    pass = false;
    std::string lot_id;
    std::string wafer_id;
};

/** ウェハ情報 */
struct WaferInfo {
    std::string lot_id;
    std::string wafer_id;
    int         slot_no = 0;
    AlignmentResult align;
    InspectionResult inspect;
    std::chrono::system_clock::time_point load_time;
    std::chrono::system_clock::time_point unload_time;
};

/** 露光レシピ */
struct ExposureRecipe {
    std::string recipe_name;
    double  energy_mj    = 25.0;   ///< 露光エネルギー [mJ/cm²]
    int     exp_time_ms  = 100;    ///< 露光時間 [ms]
    double  focus_offset = 0.0;    ///< フォーカスオフセット [μm]
    double  na           = 0.85;   ///< 開口数
    bool    immersion    = false;  ///< 液浸フラグ
};

// ============================================================
// ロガー
// ============================================================
class Logger {
public:
    enum class Level { DEBUG, INFO, WARN, ERROR };

    static Logger& getInstance() {
        static Logger instance;
        return instance;
    }

    void log(Level lvl, const std::string& msg) {
        std::lock_guard<std::mutex> lk(mtx_);
        auto now = std::chrono::system_clock::now();
        auto t   = std::chrono::system_clock::to_time_t(now);
        char buf[32];
        std::strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", std::localtime(&t));

        const char* tag = (lvl==Level::DEBUG)?"[DBG]":
                          (lvl==Level::INFO )?"[INF]":
                          (lvl==Level::WARN )?"[WRN]":"[ERR]";

        std::ostringstream oss;
        oss << buf << " " << tag << " " << msg;

        std::cout << oss.str() << "\n";
        if (file_.is_open()) file_ << oss.str() << "\n";
    }

    void openFile(const std::string& path) {
        file_.open(path, std::ios::app);
    }

private:
    Logger() = default;
    std::mutex   mtx_;
    std::ofstream file_;
};

#define LOG_I(msg) Logger::getInstance().log(Logger::Level::INFO,  msg)
#define LOG_W(msg) Logger::getInstance().log(Logger::Level::WARN,  msg)
#define LOG_E(msg) Logger::getInstance().log(Logger::Level::ERROR, msg)
#define LOG_D(msg) Logger::getInstance().log(Logger::Level::DEBUG, msg)

// ============================================================
// PLC通信クラス (SLMP / 3E フレーム)
// ============================================================
class PLCComm {
public:
    PLCComm(const std::string& ip, uint16_t port)
        : ip_(ip), port_(port), sock_(INVALID_SOCKET)
    {
#ifdef _WIN32
        WSADATA ws;
        WSAStartup(MAKEWORD(2,2), &ws);
#endif
    }

    ~PLCComm() {
        if (sock_ != INVALID_SOCKET) closesocket(sock_);
#ifdef _WIN32
        WSACleanup();
#endif
    }

    bool connect() {
        sock_ = socket(AF_INET, SOCK_DGRAM, 0);
        if (sock_ == INVALID_SOCKET) return false;

        // タイムアウト設定
#ifdef _WIN32
        DWORD tv = Const::COMM_TIMEOUT_MS;
        setsockopt(sock_, SOL_SOCKET, SO_RCVTIMEO, (char*)&tv, sizeof(tv));
#else
        timeval tv{ 0, Const::COMM_TIMEOUT_MS * 1000 };
        setsockopt(sock_, SOL_SOCKET, SO_RCVTIMEO, &tv, sizeof(tv));
#endif
        memset(&addr_, 0, sizeof(addr_));
        addr_.sin_family = AF_INET;
        addr_.sin_port   = htons(port_);
        inet_pton(AF_INET, ip_.c_str(), &addr_.sin_addr);

        LOG_I("PLC接続完了: " + ip_ + ":" + std::to_string(port_));
        return true;
    }

    /**
     * @brief データレジスタ書き込み (SLMP 3Eフレーム ワード書き込み)
     * @param start_addr 先頭デバイスアドレス (D レジスタ)
     * @param values     書き込みデータ配列
     */
    bool writeDataReg(uint16_t start_addr, const std::vector<int16_t>& values) {
        // SLMP 3Eフレーム構築
        std::vector<uint8_t> frame;
        buildSLMP3EWrite(frame, 0x1401, start_addr, 0xA8/*D*/, values);

        if (sendto(sock_, reinterpret_cast<char*>(frame.data()),
                   (int)frame.size(), 0,
                   reinterpret_cast<sockaddr*>(&addr_), sizeof(addr_)) == SOCKET_ERROR) {
            LOG_E("PLC書き込み送信失敗 D" + std::to_string(start_addr));
            return false;
        }

        // 応答受信
        uint8_t resp[64];
        int rlen = recvfrom(sock_, reinterpret_cast<char*>(resp),
                            sizeof(resp), 0, nullptr, nullptr);
        if (rlen < 0) {
            LOG_E("PLC書き込み応答タイムアウト");
            return false;
        }

        // 終了コード確認 (バイト9-10が 0x0000 で正常)
        uint16_t endCode = (uint16_t)(resp[9] | (resp[10] << 8));
        return (endCode == 0x0000);
    }

    /**
     * @brief ビットデバイス書き込み (Mコイル)
     */
    bool writeBit(uint16_t addr, bool value) {
        std::vector<uint8_t> frame;
        buildSLMP3EBitWrite(frame, addr, 0x91/*M*/, value ? 1 : 0);

        if (sendto(sock_, reinterpret_cast<char*>(frame.data()),
                   (int)frame.size(), 0,
                   reinterpret_cast<sockaddr*>(&addr_), sizeof(addr_)) == SOCKET_ERROR) {
            LOG_E("PLCビット書き込み失敗 M" + std::to_string(addr));
            return false;
        }

        uint8_t resp[64];
        int rlen = recvfrom(sock_, reinterpret_cast<char*>(resp),
                            sizeof(resp), 0, nullptr, nullptr);
        return (rlen > 0);
    }

    /**
     * @brief データレジスタ読み込み
     */
    bool readDataReg(uint16_t start_addr, int count, std::vector<int16_t>& out) {
        std::vector<uint8_t> frame;
        buildSLMP3ERead(frame, 0x0401, start_addr, 0xA8, (uint16_t)count);

        if (sendto(sock_, reinterpret_cast<char*>(frame.data()),
                   (int)frame.size(), 0,
                   reinterpret_cast<sockaddr*>(&addr_), sizeof(addr_)) == SOCKET_ERROR) {
            return false;
        }

        uint8_t resp[256];
        int rlen = recvfrom(sock_, reinterpret_cast<char*>(resp),
                            sizeof(resp), 0, nullptr, nullptr);
        if (rlen < 11 + count * 2) return false;

        out.resize(count);
        for (int i = 0; i < count; ++i) {
            out[i] = (int16_t)(resp[11 + i*2] | (resp[12 + i*2] << 8));
        }
        return true;
    }

    /**
     * @brief ビットデバイス読み込み
     */
    bool readBit(uint16_t addr, bool& out) {
        // 省略: readDataRegと同様のSLMPフレーム(コマンド0x0401,サブ0x0001)
        // 実装は同パターンにつきここでは擬似実装
        out = false;
        return true;
    }

private:
    std::string   ip_;
    uint16_t      port_;
    SOCKET        sock_;
    sockaddr_in   addr_;

    /** SLMP 3Eフレーム ワード書き込み構築 */
    void buildSLMP3EWrite(std::vector<uint8_t>& f,
                          uint16_t cmd,
                          uint16_t addr,
                          uint8_t  devCode,
                          const std::vector<int16_t>& data)
    {
        uint16_t dataLen = (uint16_t)(data.size() * 2);
        uint16_t bodyLen = 12 + dataLen;  // コマンド以降のデータ長

        f.clear();
        // サブヘッダ
        f.push_back(0x50); f.push_back(0x00); // 3Eフレーム
        // ネットワーク番号, PC番号, 要求先ユニットI/O, 要求先ユニット局番
        f.push_back(0x00); f.push_back(0xFF);
        f.push_back(0xFF); f.push_back(0x03);
        f.push_back(0x00);
        // データ長 (リトルエンディアン)
        f.push_back(bodyLen & 0xFF); f.push_back((bodyLen >> 8) & 0xFF);
        // CPU監視タイマ
        f.push_back(0x10); f.push_back(0x00);
        // コマンド・サブコマンド
        f.push_back(cmd & 0xFF); f.push_back((cmd >> 8) & 0xFF);
        f.push_back(0x00); f.push_back(0x00);
        // 先頭デバイス番号 (3バイト)
        f.push_back(addr & 0xFF); f.push_back((addr >> 8) & 0xFF); f.push_back(0x00);
        // デバイスコード
        f.push_back(devCode);
        // 書き込み点数
        f.push_back((uint8_t)data.size()); f.push_back(0x00);
        // データ
        for (auto v : data) {
            f.push_back(v & 0xFF);
            f.push_back((v >> 8) & 0xFF);
        }
    }

    void buildSLMP3EBitWrite(std::vector<uint8_t>& f,
                              uint16_t addr, uint8_t devCode, uint8_t val)
    {
        std::vector<int16_t> d = { (int16_t)val };
        buildSLMP3EWrite(f, 0x1401, addr, devCode, d);
    }

    void buildSLMP3ERead(std::vector<uint8_t>& f,
                         uint16_t cmd, uint16_t addr,
                         uint8_t devCode, uint16_t count)
    {
        f.clear();
        f.push_back(0x50); f.push_back(0x00);
        f.push_back(0x00); f.push_back(0xFF);
        f.push_back(0xFF); f.push_back(0x03); f.push_back(0x00);
        // データ長 = コマンド(2)+サブ(2)+デバイス(3+1)+点数(2) = 10
        f.push_back(0x0C); f.push_back(0x00);
        f.push_back(0x10); f.push_back(0x00);
        f.push_back(cmd & 0xFF); f.push_back((cmd >> 8) & 0xFF);
        f.push_back(0x00); f.push_back(0x00);
        f.push_back(addr & 0xFF); f.push_back((addr >> 8) & 0xFF); f.push_back(0x00);
        f.push_back(devCode);
        f.push_back(count & 0xFF); f.push_back((count >> 8) & 0xFF);
    }
};

// ============================================================
// アライメント計測クラス (カメラ画像処理)
// ============================================================
class AlignmentCamera {
public:
    /**
     * @brief ウェハアライメントマーク計測
     *
     * 実機では OpenCV / 専用SDK を使用。
     * ここでは計測ロジックの骨格を示す。
     *
     * @return AlignmentResult 計測結果
     */
    AlignmentResult measure() {
        AlignmentResult result;
        result.timestamp = std::chrono::system_clock::now();

        try {
            // --- 擬似実装 (実機では画像取得→テンプレートマッチング) ---
            // cv::Mat frame = camera_.capture();
            // auto marks = detectAlignmentMarks(frame);
            // result = calcMisalignment(marks);

            // ====== 以下は実際のアルゴリズム骨格 ======
            // 1. アライメントマーク検出 (テンプレートマッチング or ハフ円検出)
            auto markPositions = detectMarks();

            // 2. 理想位置との差分計算
            result.dx_um  = calcDx(markPositions);   // [μm]
            result.dy_um  = calcDy(markPositions);   // [μm]
            result.dtheta = calcDtheta(markPositions); // [deg]

            result.valid = true;
            LOG_I("アライメント計測: dx=" + fmt(result.dx_um) +
                  "μm dy=" + fmt(result.dy_um) +
                  "μm dθ=" + fmt(result.dtheta) + "deg");
        }
        catch (const std::exception& e) {
            LOG_E("カメラ計測エラー: " + std::string(e.what()));
            result.valid = false;
        }

        return result;
    }

private:
    struct MarkPos { double x, y; };

    /** アライメントマーク検出 (実機ではOpenCV実装) */
    std::vector<MarkPos> detectMarks() {
        // 擬似: ランダムな微小ズレを返す
        // 実際: テンプレートマッチングで4点のアライメントマーク位置を検出
        static std::mt19937 rng(std::random_device{}());
        std::normal_distribution<double> noise(0.0, 3.0); // σ=3μm

        // 理想位置 (4点: mm→pixel変換係数込み)
        std::vector<MarkPos> ideal = {
            {-20000.0, -20000.0},
            { 20000.0, -20000.0},
            { 20000.0,  20000.0},
            {-20000.0,  20000.0}
        };

        std::vector<MarkPos> detected;
        for (auto& p : ideal) {
            detected.push_back({ p.x + noise(rng), p.y + noise(rng) });
        }
        return detected;
    }

    double calcDx(const std::vector<MarkPos>& marks) {
        // 4点の重心X差分 → [μm] 変換
        double sumDx = 0;
        std::vector<double> idealX = {-20000, 20000, 20000, -20000};
        for (size_t i = 0; i < marks.size(); ++i)
            sumDx += marks[i].x - idealX[i];
        // pixel→μm変換係数 (例: 0.1μm/pixel)
        return (sumDx / marks.size()) * 0.1;
    }

    double calcDy(const std::vector<MarkPos>& marks) {
        double sumDy = 0;
        std::vector<double> idealY = {-20000, -20000, 20000, 20000};
        for (size_t i = 0; i < marks.size(); ++i)
            sumDy += marks[i].y - idealY[i];
        return (sumDy / marks.size()) * 0.1;
    }

    double calcDtheta(const std::vector<MarkPos>& marks) {
        // 対角マーク間ベクトルの角度差から回転ズレを計算
        if (marks.size() < 2) return 0.0;
        double vx = marks[1].x - marks[0].x;
        double vy = marks[1].y - marks[0].y;
        double angle = std::atan2(vy, vx) * 180.0 / M_PI;
        double idealAngle = 0.0; // 理想角度
        return angle - idealAngle;
    }

    std::string fmt(double v) {
        std::ostringstream os;
        os << std::fixed << std::setprecision(3) << v;
        return os.str();
    }

    // cv::VideoCapture camera_; // 実機ではOpenCVカメラオブジェクト
};

// ============================================================
// 露光制御クラス
// ============================================================
class ExposureUnit {
public:
    explicit ExposureUnit(PLCComm& plc) : plc_(plc) {}

    /**
     * @brief 露光実行
     * @param recipe 露光レシピ
     */
    bool expose(const ExposureRecipe& recipe) {
        LOG_I("露光開始: " + recipe.recipe_name +
              " energy=" + std::to_string(recipe.energy_mj) + "mJ/cm²");

        // 露光時間をPLCに設定
        std::vector<int16_t> params = { (int16_t)recipe.exp_time_ms };
        if (!plc_.writeDataReg(Const::D_EXP_TIME, params)) {
            LOG_E("露光時間設定失敗");
            return false;
        }

        // シャッター開→露光→シャッター閉 はPLCラダーで制御
        // ここでは完了待ち
        auto start = std::chrono::steady_clock::now();
        bool done = false;

        while (!done) {
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
            auto elapsed = std::chrono::steady_clock::now() - start;
            if (elapsed > std::chrono::milliseconds(recipe.exp_time_ms + 500)) {
                LOG_E("露光タイムアウト");
                return false;
            }

            // PLCからステップ確認
            std::vector<int16_t> step;
            if (plc_.readDataReg(Const::D_STEP, 1, step) && step[0] == 7) {
                done = true; // ステップ7=検査ステップに進んだ
            }
        }

        LOG_I("露光完了");
        return true;
    }

private:
    PLCComm& plc_;
};

// ============================================================
// 検査クラス
// ============================================================
class InspectionUnit {
public:
    /**
     * @brief 露光後検査
     * @param wafer  ウェハ情報
     * @return       検査結果
     */
    InspectionResult inspect(const WaferInfo& wafer) {
        InspectionResult result;
        result.lot_id   = wafer.lot_id;
        result.wafer_id = wafer.wafer_id;

        LOG_I("検査開始: LOT=" + wafer.lot_id + " WFR=" + wafer.wafer_id);

        // --- CD（線幅）計測 ---
        result.cd_error_nm = measureCD();

        // --- 欠陥検査 ---
        detectDefects(result);

        // --- 判定 ---
        if (std::abs(result.cd_error_nm) > Const::INSPECT_CD_TOL) {
            result.verdict = InspectionResult::Verdict::NG_CD;
            result.pass    = false;
            LOG_W("CD誤差NG: " + std::to_string(result.cd_error_nm) + "nm");
        }
        else if (result.defect_ratio > Const::INSPECT_DEF_MAX) {
            result.verdict = InspectionResult::Verdict::NG_DEFECT;
            result.pass    = false;
            LOG_W("欠陥NGl: ratio=" + std::to_string(result.defect_ratio) + "%");
        }
        else {
            result.verdict = InspectionResult::Verdict::OK;
            result.pass    = true;
            LOG_I("検査OK");
        }

        return result;
    }

private:
    /** CD計測 (SEM/光学計測シミュレーション) */
    double measureCD() {
        // 実機: SEMまたは散乱光計測器からの読み値
        static std::mt19937 rng(42);
        std::normal_distribution<double> noise(0.0, 0.8); // σ=0.8nm
        return noise(rng);
    }

    /** 欠陥検査 */
    void detectDefects(InspectionResult& result) {
        // 実機: 暗視野検査カメラ画像の差分解析
        static std::mt19937 rng(123);
        std::uniform_real_distribution<double> dist(0.0, 0.15);
        result.defect_ratio = dist(rng);
        result.defect_count = (int)(result.defect_ratio * 100);
    }
};

// ============================================================
// メイン制御クラス
// ============================================================
class ExposureController {
public:
    ExposureController()
        : plc_(Const::PLC_IP, Const::PLC_PORT)
        , expUnit_(plc_)
        , running_(false)
        , estopDetected_(false)
    {}

    bool init() {
        Logger::getInstance().openFile("exposure_ctrl.log");
        LOG_I("=== LSI露光装置制御システム 起動 ===");

        if (!plc_.connect()) {
            LOG_E("PLC接続失敗");
            return false;
        }

        // レシピ読み込み
        recipe_.recipe_name = "ArF_28nm_STD";
        recipe_.energy_mj   = Const::EXPOSURE_ENERGY;
        recipe_.exp_time_ms = 100;
        recipe_.na          = 0.85;

        LOG_I("レシピ設定完了: " + recipe_.recipe_name);
        return true;
    }

    /** メインループ起動 */
    void run() {
        running_ = true;

        // 非常停止監視スレッド
        std::thread estopThread([this]() { monitorEstop(); });

        // メイン制御ループ
        while (running_) {
            try {
                mainLoop();
            }
            catch (const std::exception& e) {
                LOG_E("制御ループ例外: " + std::string(e.what()));
                std::this_thread::sleep_for(std::chrono::milliseconds(500));
            }
        }

        estopThread.join();
    }

    void stop() { running_ = false; }

private:
    PLCComm           plc_;
    AlignmentCamera   camera_;
    ExposureUnit      expUnit_;
    InspectionUnit    inspUnit_;
    ExposureRecipe    recipe_;
    std::atomic<bool> running_;
    std::atomic<bool> estopDetected_;

    // 統計
    int okCount_ = 0;
    int ngCount_ = 0;

    /** 非常停止監視 */
    void monitorEstop() {
        while (running_) {
            bool estop = false;
            plc_.readBit(Const::M_ESTOP, estop);
            if (estop && !estopDetected_) {
                estopDetected_ = true;
                LOG_E("非常停止検出！");
                running_ = false;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
        }
    }

    /** メイン制御ループ (PLCステップと協調) */
    void mainLoop() {
        std::vector<int16_t> stepVec;
        if (!plc_.readDataReg(Const::D_STEP, 1, stepVec)) {
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
            return;
        }
        int step = stepVec[0];

        switch (step) {
            case 3:  handleAlignment();  break;
            case 6:  handleExposure();   break;
            case 7:  handleInspection(); break;
            default:
                std::this_thread::sleep_for(std::chrono::milliseconds(20));
                break;
        }
    }

    /** ステップ3: アライメント処理 */
    void handleAlignment() {
        // アライメントトリガ確認
        bool trigActive = false;
        plc_.readBit(Const::M_ALIGN_TRIG, trigActive);
        if (!trigActive) return;

        LOG_I("--- アライメント処理開始 ---");

        for (int retry = 0; retry < Const::ALIGN_RETRY; ++retry) {
            AlignmentResult ar = camera_.measure();

            if (!ar.valid) {
                LOG_W("計測無効 リトライ " + std::to_string(retry+1));
                continue;
            }

            // ズレ量をPLCに書き込み (μm*10 → int16_t)
            std::vector<int16_t> misalign = {
                (int16_t)std::round(ar.dx_um   * 10.0),
                (int16_t)std::round(ar.dy_um   * 10.0),
                (int16_t)std::round(ar.dtheta  * 1000.0)
            };
            plc_.writeDataReg(Const::D_MISALIGN_X, misalign);

            // 補正範囲チェック
            if (std::abs(ar.dx_um) > Const::ALIGN_MAX_XY ||
                std::abs(ar.dy_um) > Const::ALIGN_MAX_XY) {
                LOG_E("補正範囲超過: dx=" + std::to_string(ar.dx_um) +
                      " dy=" + std::to_string(ar.dy_um));
                plc_.writeBit(Const::M_NG_FLAG, true);
                return;
            }

            // 許容範囲内なら補正完了シグナル
            if (std::abs(ar.dx_um) < Const::ALIGN_TOL_XY &&
                std::abs(ar.dy_um) < Const::ALIGN_TOL_XY &&
                std::abs(ar.dtheta) < Const::ALIGN_TOL_TH) {
                LOG_I("アライメント完了 残留誤差: dx=" +
                      std::to_string(ar.dx_um) + "μm");
                plc_.writeBit(Const::M_ALIGN_TRIG, false); // トリガリセット
                return;
            }

            // PLCが補正を実行するのを待つ
            std::this_thread::sleep_for(std::chrono::milliseconds(200));
        }

        LOG_E("アライメントリトライ上限超過 → NG");
        plc_.writeBit(Const::M_NG_FLAG, true);
    }

    /** ステップ6: 露光処理 */
    void handleExposure() {
        LOG_I("--- 露光処理開始 ---");
        if (!expUnit_.expose(recipe_)) {
            LOG_E("露光失敗");
            plc_.writeBit(Const::M_NG_FLAG, true);
        }
    }

    /** ステップ7: 検査処理 */
    void handleInspection() {
        LOG_I("--- 検査処理開始 ---");

        WaferInfo wafer;
        wafer.lot_id   = generateLotId();
        wafer.wafer_id = generateWaferId();

        InspectionResult ir = inspUnit_.inspect(wafer);

        if (ir.pass) {
            plc_.writeBit(Const::M_OK_FLAG, true);
            plc_.writeBit(Const::M_NG_FLAG, false);
            ++okCount_;
        } else {
            plc_.writeBit(Const::M_NG_FLAG, true);
            plc_.writeBit(Const::M_OK_FLAG, false);
            ++ngCount_;
        }

        // 統計ログ
        int total = okCount_ + ngCount_;
        double yield = total > 0 ? (okCount_ * 100.0 / total) : 0.0;
        LOG_I("生産統計: OK=" + std::to_string(okCount_) +
              " NG=" + std::to_string(ngCount_) +
              " 歩留=" + std::to_string(yield) + "%");

        // 検査結果をPLC D レジスタに反映
        std::vector<int16_t> counts = {
            (int16_t)ngCount_,
            (int16_t)okCount_
        };
        plc_.writeDataReg(Const::D_NG_COUNT, counts);
    }

    std::string generateLotId() {
        auto t = std::chrono::system_clock::to_time_t(
                     std::chrono::system_clock::now());
        char buf[16];
        std::strftime(buf, sizeof(buf), "LOT%m%d%H%M", std::localtime(&t));
        return buf;
    }

    std::string generateWaferId() {
        static int wno = 1;
        std::ostringstream os;
        os << "W" << std::setw(3) << std::setfill('0') << wno++;
        return os.str();
    }
};

// ============================================================
// エントリポイント
// ============================================================
int main() {
    ExposureController ctrl;

    if (!ctrl.init()) {
        std::cerr << "初期化失敗\n";
        return 1;
    }

    ctrl.run();
    return 0;
}