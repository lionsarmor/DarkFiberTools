#!/usr/bin/ruby
require 'logger'
require 'json' # Ensure JSON module is available

$:.push File.dirname(__FILE__)
require 'read'
require 'dump'

# ============== main ===========================
if __FILE__ == $0
  if ARGV.length < 2 then
    puts "USAGE: #{__FILE__} SOR-file output-directory"
    exit
  end

  otdrfile = ARGV[0]
  output_dir = ARGV[1]

  # Ensure the output directory exists
  unless Dir.exist?(output_dir)
    Dir.mkdir(output_dir)
  end

  $logger = Logger.new(STDOUT)
  $logger.formatter = proc { |severity, datetime, progname, msg| "#{severity}: #{msg}\n" }

  sorparse = SORparse.new(otdrfile)
  results = {}
  trace = []

  begin
    sorparse.run(results, trace, debug=false)

    # Write results to JSON file in the specified output directory
    resultsfile = File.join(output_dir, File.basename(otdrfile, ".*") + "-dump.json")
    Dump::jsonfile(results, resultsfile)

    # Comment out trace file generation
    # tracefile = File.join(output_dir, File.basename(otdrfile, ".*") + "-trace.dat")
    # Dump::tracefile(trace, tracefile)

  rescue => e
    $logger.error("Error processing file: #{e.message}")
    exit(1)
  end
end
