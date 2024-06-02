<?php

/**
 * Bundled Service XLS Merger
 *
 * @package     BundledServiceXLSMerger
 * @author      Henri Susanto
 * @copyright   2024 Henri Susanto
 * @license     GPL-2.0-or-later
 *
 * @wordpress-plugin
 * Plugin Name: Bundled Service XLS Merger
 * Plugin URI:  https://github.com/susantohenri/bundled-service-xls-merger
 * Description: Merge Uploaded RFP Files for Bundled Service
 * Version:     1.0.0
 * Author:      Henri Susanto
 * Author URI:  https://github.com/susantohenri/
 * Text Domain: BundledServiceXLSMerger
 * License:     GPL v2 or later
 * License URI: http://www.gnu.org/licenses/gpl-2.0.txt
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

define('bundled_service_xls_merger_form_id', 58);

add_action('wp_ajax_nopriv_rbundle_custom_submit_dropzone', 'FrmProFieldsController::ajax_upload');
add_action('wp_ajax_rbundle_custom_submit_dropzone', 'FrmProFieldsController::ajax_upload');
add_action('frm_after_create_entry', 'bundled_service_xls_merger', 30, 2);
add_action('frm_after_update_entry', 'bundled_service_xls_merger', 10, 2);

function bundled_service_xls_merger($entry_id, $form_id)
{
    if (bundled_service_xls_merger_form_id != $form_id) return true;
    if (!isset($_POST['bundled_children'])) return true;
    if (!isset($_POST['bundled_children']['3713'])) return true;

    global $wpdb;
    $final_file = new Spreadsheet();
    $is_multi_bus = 1 < count($_POST['business_name']);
    $media_path = wp_get_upload_dir()['basedir'] . "/formidable/{$form_id}/";
    $media_url = site_url() . "/wp-content/uploads/formidable/{$form_id}/";

    foreach ($_POST['bundled_children']['3713'] as $media_id_sheet_name) {
        if (!strpos($media_id_sheet_name, '|')) continue;

        $media_id_sheet_name = explode('|', $media_id_sheet_name);
        $media_id = $media_id_sheet_name[0];
        $sheet_name = $media_id_sheet_name[1];

        $guid = $wpdb->get_var("SELECT guid FROM {$wpdb->prefix}posts WHERE ID = {$media_id}");
        $file_path = explode('/', $guid);
        $file_path = end($file_path);
        $file_path = $media_path . $file_path;
        if (!file_exists($file_path)) continue;

        $service_file = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
        $service_file->getActiveSheet()->setTitle($sheet_name);
        $final_file->addExternalSheet($service_file->getActiveSheet());

        unlink($file_path);
        $wpdb->delete("{$wpdb->prefix}posts", ['ID' => $media_id]);
    }

    $final_file->removeSheetByIndex(0);
    $writer = new Xlsx($final_file);
    $answer_5356 = $_POST['item_meta'][5356];
    $date = date('d-m-Y');

    $final_file_name = 'Rbundle RFP';
    $final_file_name .= $is_multi_bus ? ' - Business Bundle ' : ' - Service Bundle ';
    $final_file_name .= "{$date} - {$answer_5356}";
    $final_file_path = "{$media_path}{$final_file_name}.xlsx";
    $final_file_url = "{$media_url}{$final_file_name}.xlsx";

    $writer->save($final_file_path);
    $final_media_id = wp_insert_post([
        'guid' => $final_file_url,
        'post_type' => 'attachment',
        'post_mime_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ]);

    $answer_id_5361 = $wpdb->get_var("SELECT id FROM {$wpdb->prefix}frm_item_metas WHERE item_id = {$entry_id} AND field_id = 5361");
    if ($answer_id_5361) $wpdb->update("{$wpdb->prefix}frm_item_metas", ['meta_value' => $final_media_id], ['id' => $answer_id_5361]);
    else $wpdb->insert("{$wpdb->prefix}frm_item_metas", [
        'meta_value' => $final_media_id,
        'field_id' => 5361,
        'item_id' => $entry_id
    ]);
}